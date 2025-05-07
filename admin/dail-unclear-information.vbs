'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL UNCLEAR INFORMATION.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 0
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
call changelog_update("02/12/2025", "Fixed issue where navigation to ELIG/FS used DAIL message month instead of current month.", "Mark Riegel, Hennepin County")
call changelog_update("01/13/2025", "Improved DAIL navigation handling.", "Mark Riegel, Hennepin County")
call changelog_update("08/21/2023", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------FUNCTIONS
'Add function for X Number restart functionality
Function create_array_of_all_active_x_numbers_in_county_with_restart(array_name, two_digit_county_code, restart_status, restart_x_number)
'--- This function is used to grab all active X numbers in a county
'~~~~~ array_name: name of array that will contain all the x numbers
'~~~~~ county_code: inserted by reading the county code under REPT/USER
'===== Keywords: MAXIS, array, worker number, create
	'Getting to REPT/USER
	Call navigate_to_MAXIS_screen("REPT", "USER")
	PF5 'Hitting PF5 to force sorting, which allows directly selecting a county
	Call write_value_and_transmit(county_code, 21, 6)  	'Inserting county

	MAXIS_row = 7  'Declaring the MAXIS row
	array_name = ""    'Blanking out array_name in case this has been used already in the script

    Found_restart_worker = False    'defaulting to false. Will become true when the X number is found.
	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
            If restart_status = True then
                If trim(UCase(worker_ID)) = trim(UCase(restart_x_number)) then
                    Found_restart_worker = True
                End if
                If Found_restart_worker = True then array_name = trim(array_name & " " & worker_ID)				'writing to variable
            Else
                array_name = trim(array_name & " " & worker_ID)				'writing to variable
            End if
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then
			PF8			'getting to next screen
			MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank

    array_name = split(array_name)
End function

Function check_and_add_new_jobs_panel(testing_status)
    'Need to navigate to JOBS panel for CM if not there already so will check if at CM right now
    EMReadScreen JOBS_footer_month_and_year, 5, 20, 55

    If JOBS_footer_month_and_year <> CM_mo & " " & CM_yr then 
        If testing_status = True Then MsgBox "Testing -- Need to navigate to CM"
        'PF3 back to DAIL and navigate to CASE/CURR to change the footer month and get to JOBS panel for CM
        PF3
        Call write_value_and_transmit("H", dail_row, 3)
        EMReadScreen curr_panel_check, 4, 2, 55
        If curr_panel_check <> "CURR" Then MsgBox "Testing -- not at CASE/CURR"
        EMWriteScreen "STAT", 20, 22
        EMWriteScreen CM_mo, 20, 54
        EMWriteScreen CM_yr, 20, 57
        Call write_value_and_transmit("JOBS", 20, 69)

        'Open the first JOBS panel of the caregiver reference number
        EMWriteScreen HIRE_memb_number, 20, 76
        Call write_value_and_transmit("01", 20, 79)
    Else    
        If testing_status = True Then MsgBox "Testing -- Already at CM JOBS panel"
    End If

    'Ensure we are on JOBS panel
    EmReadScreen jobs_panel_nav_check, 4, 2, 45
    If jobs_panel_nav_check <> "JOBS" Then MsgBox "Testing -- Not on JOBS panel. Stop here"
    
    If testing_status = True Then MsgBox "Testing -- Ensure that we are on correct HH Member. Should be at HH Member: " & HIRE_memb_number

    'Check if no JOBS panel exists on HH Memb JOBS panel
    EmReadScreen jobs_panel_check, 40, 24, 2

    If InStr(jobs_panel_check, "DOES NOT EXIST") Then
        'There are no JOBS panels for this HH member. The script will add a new JOBS panel for the member
        If testing_status = True Then MsgBox "Testing -- No JOBS panel exist. Script will create new panel and fill it out. STOP HERE if needed in production."

        Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

        'Validation to ensure that script is able to open a new JOBS panel
        EmReadScreen panel_count_plus_one_check, 1, 2, 73
        panel_count_plus_one_check = panel_count_plus_one_check * 1
        EmReadScreen panel_count_total_check, 1, 2, 78
        panel_count_total_check = panel_count_total_check * 1

        If panel_count_plus_one_check <> panel_count_total_check + 1 then 
            If testing_status = True Then MsgBox "Testing -- unable to open a new JOBS panel. Will note in spreadsheet and continue"
            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "MAXIS programs are inactive. Unable to add a new JOBS panel for M" & HIRE_memb_number & ". Review needed." & " Message should not be deleted."
        Else
            
            If testing_status = True Then MsgBox "Testing -- Script opened JOBS panel. Will add new panel"

            'Reads footer month for updating the panel
            EMReadScreen JOBS_footer_month, 2, 20, 55	
            EMReadScreen JOBS_footer_year, 2, 20, 58	

            'Write the date hired date from NDNH message to JOBS panel
            Call create_MAXIS_friendly_date(date_hired, 0, 9, 35)

            'Writes information to JOBS panel
            EMWriteScreen "W", 5, 34
            EMWriteScreen "4", 6, 34
            EMWriteScreen HIRE_employer_name, 7, 42
                
            IF month_hired = JOBS_footer_month and year_hired = JOBS_footer_year THEN
                'If the footer month on the JOBS panel matches the month from the HIRE message then it writes the actual hired date from the message to the panel
                Call create_MAXIS_friendly_date(date_hired, 0, 12, 54)
            ELSE
                'Otherwise, write the panel footer month and date to the new panel
                EmWriteScreen JOBS_footer_month, 12, 54
                EMWriteScreen "01", 12, 57
                EmWriteScreen JOBS_footer_year, 12, 60
            END IF

            'Puts $0 in as the received income amt and 0 worked hours
            EMWriteScreen "0", 12, 67				
            EMWriteScreen "0", 18, 72	
            
            If testing_status = True Then msgbox "Testing -- Review the JOBS panel. Any potential errors or issues before it continues?"
            
            'Opens FS PIC
            Call write_value_and_transmit("X", 19, 38)
                
            'Write today's date to calculation since added today
            Call create_MAXIS_friendly_date(date, 0, 5, 34)
            
            'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
            EMWriteScreen "1", 5, 64
            EMWriteScreen "0", 8, 64
            EMWriteScreen "0", 9, 66
            If testing_status = True Then msgbox "Testing -- Review the PIC panel. Any potential errors or issues before it continues?"

            transmit
            EmReadScreen PIC_warning, 7, 20, 6
            IF PIC_warning = "WARNING" then transmit 'to clear message
            transmit 'back to JOBS panel
            If testing_status = True Then Msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
            transmit 'to save JOBS panel

            'Check if information is expiring and needs to be added to a future month
            EMReadScreen expired_check, 6, 24, 17 
            EMReadScreen data_expiration_month, 2, 24, 27
            EMReadScreen jobs_panel_month, 2, 20, 55

            If expired_check = "EXPIRE" THEN 
                Do
                    'Do loop to add JOBS panels to every month from DAIL month through CM
                    If testing_status = True Then msgbox "Testing -- New JOBS panel is expiring so it needs to be added to CM + 1 as well"

                    'PF3 to go to STAT/WRAP
                    PF3

                    'Check to make sure on STAT/WRAP
                    EMReadScreen stat_wrap_check, 19, 2, 32
                    If Instr(stat_wrap_check, "Wrap") = 0 Then MsgBox "Testing -- It didn't go to STAT/WRAP for some reason. Stop here!!"

                    'Build do loop to get to expiration month so that it isn't creating a bunch of duplicate JOBS panels
                    If data_expiration_month <> jobs_panel_month Then
                        If testing_status = True Then msgbox "Testing -- JOBS panel expires in future month"
                        Do
                            Call write_value_and_transmit("Y", 16, 54)
                            EMReadScreen stat_wrap_month_check, 2, 20, 55
                            If stat_wrap_month_check = data_expiration_month Then
                                'Script has reached the expiration month, it will go to next month and then exit
                                If testing_status = True Then msgbox "Testing -- script has found matching month"
                                PF3
                                Call write_value_and_transmit("Y", 16, 54)
                                Exit Do
                            Else
                                'Script has not yet reached the expiration month, it will PF3 back to STAT/WRAP to move to next month
                                If testing_status = True Then msgbox "Testing -- script has not reached matching month"
                                PF3
                            End If
                        Loop
                    Else
                        'If the expiration month and the jobs panel month are the same then it should add to next month too since it will expire at end of month
                        Call write_value_and_transmit("Y", 16, 54)
                    End If

                    'Navigate to STAT/JOBS
                    Call write_value_and_transmit("JOBS", 20, 71)

                    EMReadScreen jobs_panel_nav_check, 8, 2, 43
                    If InStr(jobs_panel_nav_check, "JOBS") = 0 Then MsgBox "Testing -- Stop here. Not at JOBS panel"

                    If testing_status = True Then MsgBox "Testing -- Is it at the month after expiration? Expiration month was " & data_expiration_month

                    'Navigate to HH member
                    Call write_value_and_transmit(HIRE_memb_number, 20, 76)

                    'Making sure there aren't 5 jobs already
                    EMReadScreen five_jobs_check, 1, 2, 78
                    
                    If five_jobs_check = "5" Then 
                        script_end_procedure_with_error_report("Testing -- There are 5 JOBS panels already, it will error out. Must stop here!")
                    Else
                        Call write_value_and_transmit("NN", 20, 79)
                    End If
                    
                    EmReadScreen panel_count_plus_one_check, 1, 2, 73
                    panel_count_plus_one_check = panel_count_plus_one_check * 1
                    EmReadScreen panel_count_total_check, 1, 2, 78
                    panel_count_total_check = panel_count_total_check * 1

                    If panel_count_plus_one_check <> panel_count_total_check + 1 then script_end_procedure_with_error_report("Testing -- Unable to open a new JOBS panel. Script will stop here.")

                    'Reads footer month for updating the panel
                    EMReadScreen JOBS_footer_month, 2, 20, 55	
                    EMReadScreen JOBS_footer_year, 2, 20, 58	

                    'Write the date hired date from NDNH message to JOBS panel
                    Call create_MAXIS_friendly_date(date_hired, 0, 9, 35)

                    'Writes information to JOBS panel
                    EMWriteScreen "W", 5, 34
                    EMWriteScreen "4", 6, 34
                    EMWriteScreen HIRE_employer_name, 7, 42
                    
                    'Looking at CM + 1 so won't match the message, just writes footer month to panel
                    EmWriteScreen JOBS_footer_month, 12, 54
                    EMWriteScreen "01", 12, 57
                    EmWriteScreen JOBS_footer_year, 12, 60

                    'Puts $0 in as the received income amt
                    EMWriteScreen "0", 12, 67				
                    'Puts 0 hours in as the worked hours
                    EMWriteScreen "0", 18, 72		

                    If testing_status = True Then msgbox "Testing - Does everything look good on JOBS panel before heading to PIC?"
                    
                    'Opens FS PIC
                    Call write_value_and_transmit("X", 19, 38)
                    'Writes today's date on the panel
                    Call create_MAXIS_friendly_date(date, 0, 5, 34)

                    'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                    EMWriteScreen "1", 5, 64
                    EMWriteScreen "0", 8, 64
                    EMWriteScreen "0", 9, 66
                    If testing_status = True Then msgbox "Testing - Does everything look good on JOBS panel before saving the PIC?"
                    
                    transmit
                    EmReadScreen PIC_warning, 7, 20, 6
                    IF PIC_warning = "WARNING" then transmit 'to clear message
                    transmit 'back to JOBS panel
                    If testing_status = True Then msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                    transmit 'to save JOBS panel
                    
                    'Check if information is expiring and needs to be added to CM + 1
                    EMReadScreen expired_check, 6, 24, 17 
                    EMReadScreen data_expiration_month, 2, 24, 27
                    EMReadScreen jobs_panel_month, 2, 20, 55 
                    
                    If expired_check <> "EXPIRE" THEN
                        'If data is not expiring, then the script can exit the do loop
                        If testing_status = True Then msgbox "Testing -- No expiration date. It will exit the do loop"
                        Exit Do
                    Else
                        If testing_status = True Then msgbox "Testing -- Data is expiring. It will continue with the do loop"
                    End If

                Loop

            End If

            'Write information to CASE/NOTE
            If testing_status = True Then MsgBox "Testing -- Script will now CASE/NOTE information. Navigate to CASE/NOTE"

            'PF4 to navigate to CASE/NOTE
            PF4

            EMReadScreen jobs_panel_not_saved, 25, 24, 2
            'If unable to navigate to CASE/NOTE due to not saving JOBS panel, then another transmit is needed
            If instr(jobs_panel_not_saved, "CASE OR PERSON NOTES ARE") Then 
                transmit
                PF4
            End If

            EMReadScreen case_note_check, 4, 2, 45
            If case_note_check <> "NOTE" then MsgBox "Testing -- not at case note stop here1"

            'Open new CASE/NOTE
            PF9

            'Write information depending on whether NDNH or SDNH message
            If InStr(dail_msg, "NDNH MEMB") Then
                CALL write_variable_in_case_note("-NDNH Match for (M" & HIRE_memb_number & ") for " & trim(HIRE_employer_name) & "-")
                CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
                CALL write_variable_in_case_note("MAXIS NAME: " & NDNH_maxis_name)
                CALL write_variable_in_case_note("NEW HIRE NAME: " & NDNH_new_hire_name)
                CALL write_variable_in_case_note("EMPLOYER: " & HIRE_employer_name)
                CALL write_variable_in_case_note("---")
                CALL write_variable_in_case_note("NO CORRESPONDING JOBS PANEL EXISTED FOR EMPLOYER NOTED IN HIRE MESSAGE. STAT/JOBS PANEL ADDED FOR EMPLOYER IDENTIFIED IN HIRE DAIL MESSAGE. INFC CLEARED.")
                CALL write_variable_in_case_note("---")
                CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                CALL write_variable_in_case_note("---")
                CALL write_variable_in_case_note(worker_signature)
            ElseIf InStr(dail_msg, "SDNH NEW JOB DETAILS") Then
                CALL write_variable_in_case_note("-SDNH Match for (M" & HIRE_memb_number & ") for " & trim(HIRE_employer_name) & "-")
                CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
                CALL write_variable_in_case_note("MAXIS NAME: " & SDNH_maxis_name)
                CALL write_variable_in_case_note("NEW HIRE NAME: " & SDNH_new_hire_name)
                CALL write_variable_in_case_note("EMPLOYER: " & HIRE_employer_name)
                CALL write_variable_in_case_note("---")
                CALL write_variable_in_case_note("NO CORRESPONDING JOBS PANEL EXISTED FOR EMPLOYER NOTED IN HIRE MESSAGE. STAT/JOBS PANEL ADDED FOR EMPLOYER IDENTIFIED IN HIRE DAIL MESSAGE. HIRE MESSAGE DELETED.")
                CALL write_variable_in_case_note("---")
                CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                CALL write_variable_in_case_note("---")
                CALL write_variable_in_case_note(worker_signature)
            Else
                Msgbox "Testing -- something went wrong when writing the CASE/NOTE. Appears that message is neither NDNH or SDNH"
            End If

            If testing_status = True Then msgbox "Testing -- The script is about to save the CASE/NOTE. Stop here if in testing or production"

            'PF3 to save the CASE/NOTE
            PF3
            
            'PF3 to STAT/WRAP or JOBS
            PF3
            
            EMReadScreen panel_nav_check, 4, 2, 46
            If panel_nav_check <> "WRAP" Then
                PF3
                If testing_status = True Then msgbox "Testing -- The script should now be at STAT/WRAP. If it is not, then stop here."
            End If

            If testing_status = True Then msgbox "Testing -- No jobs panels existed. Created JOBS panel(s) through CM"
            
            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No JOBS panels exist for household member number: " & HIRE_memb_number & ". JOBS Panel and CASE/NOTE added for employer noted in HIRE message. Message should be deleted.")
        End If
    ElseIf InStr(jobs_panel_check, "NOT IN THE HOUSEHOLD") Then

        If testing_status = True then msgbox "Testing -- member is not in household for CM so will not add JOBS panel and will skip message."

        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "There does not appear to be an exactly matching JOBS panel for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & " since the HH member does not exist on the case for CM. Review needed." & " Message should not be deleted."

    Else
        'There is at least 1 JOBS panel
        If testing_status = True Then MsgBox "Testing -- there is at least 1 JOBS panel."

        'Read the employer name, but only first 20 characters to align with max length for HIRE message for NDNH messages
        EMReadScreen employer_name_jobs_panel, 20, 7, 42
        employer_name_jobs_panel = trim(replace(employer_name_jobs_panel, "_", " "))

        'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
        employer_name_jobs_panel_split = split(employer_name_jobs_panel, " ")

        If len(employer_name_jobs_panel_split(0)) < 4 and Ubound(employer_name_jobs_panel_split) > 0 Then
            employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0) & " " & employer_name_jobs_panel_split(1)
            If testing_status = True Then MsgBox "First word less than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word  
        Else
            employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0)   
            If testing_status = True Then MsgBox "First word longer than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word
        End If

        If instr(len(employer_name_jobs_panel_first_word), employer_name_jobs_panel_first_word, ",") = len(employer_name_jobs_panel_first_word) then 
            employer_name_jobs_panel_first_word = Mid(employer_name_jobs_panel_first_word, 1, len(employer_name_jobs_panel_first_word) - 1)
            If testing_status = True Then MsgBox "Last character is a comma. employer_name_jobs_panel_first_word is now " & employer_name_jobs_panel_first_word
        End If

        If employer_name_jobs_panel = HIRE_employer_name Then
            'Add here
            If testing_status = True Then msgbox "Testing -- The employer names match exactly. Will add to delete list and TIKL delete list."

            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & ". JOBS panel matches HIRE employer name exactly. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 

            list_of_TIKLs_to_delete = list_of_TIKLs_to_delete & tikl_case_number & "-" & tikl_case_name & "-" & "VERIFICATION OF " & HIRE_employer_name_TIKL & "*" 
            ' "Verification of " & employer & "job via NEW HIRE should"
            If testing_status = True Then MsgBox list_of_TIKLs_to_delete

        ElseIf employer_name_jobs_panel_first_word = HIRE_employer_name_first_word Then

            If testing_status = True Then msgbox "Testing -- there is an exact match for employer name first word only. Will add to delete list and TIKL delete list."

            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & ". JOBS panel matches first word of HIRE employer name. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 

            list_of_TIKLs_to_delete = list_of_TIKLs_to_delete & tikl_case_number & "-" & tikl_case_name & "-" & "VERIFICATION OF " & HIRE_employer_name_TIKL & "*" 
            If testing_status = True Then MsgBox list_of_TIKLs_to_delete

        Else
            'Check how many panels exist for the HH member
            EMReadScreen jobs_panels_count, 1, 2, 78
            'Convert jobs_panels_count to a number
            jobs_panels_count = jobs_panels_count * 1
            'If there is more than just 1 JOBS panel, loop through them all to check for matching employers
            If jobs_panels_count = 1 Then
                If testing_status = True Then MsgBox "Testing -- There is only one JOBS panel and they do not match. The script will skip the message since there is no exact match"

                'Set variable below to true to trigger dialog
                no_exact_JOBS_panel_matches = True
            
            ElseIf jobs_panels_count <> 1 Then
                If testing_status = True Then MsgBox "Testing -- There are multiple JOBS panels. Script will determine if there are any perfect matches."
                
                'Set incrementor for do loop
                panel_count = 1

                Do
                    panel_count = panel_count + 1
                    EMWriteScreen HIRE_memb_number, 20, 76
                    Call write_value_and_transmit("0" & panel_count, 20, 79)

                    'Read the employer name
                    EMReadScreen employer_name_jobs_panel, 20, 7, 42
                    employer_name_jobs_panel = trim(replace(employer_name_jobs_panel, "_", " "))

                    'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
                    employer_name_jobs_panel_split = split(employer_name_jobs_panel, " ")

                    If len(employer_name_jobs_panel_split(0)) < 4 and Ubound(employer_name_jobs_panel_split) > 0 Then
                        employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0) & " " & employer_name_jobs_panel_split(1)
                        If testing_status = True Then MsgBox "First word less than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word  
                    Else
                        employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0)   
                        If testing_status = True Then MsgBox "First word longer than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word
                    End If

                    If instr(len(employer_name_jobs_panel_first_word), employer_name_jobs_panel_first_word, ",") = len(employer_name_jobs_panel_first_word) then 
                        employer_name_jobs_panel_first_word = Mid(employer_name_jobs_panel_first_word, 1, len(employer_name_jobs_panel_first_word) - 1)
                        If testing_status = True Then MsgBox "Last character is a comma. employer_name_jobs_panel_first_word is now " & employer_name_jobs_panel_first_word
                    End If

                    If employer_name_jobs_panel = HIRE_employer_name Then
                        If testing_status = True Then msgbox "Testing -- The employer names match exactly. Will add to delete list and TIKL delete list."
            
                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & ". JOBS panel matches HIRE employer name exactly. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 
            
                        list_of_TIKLs_to_delete = list_of_TIKLs_to_delete & tikl_case_number & "-" & tikl_case_name & "-" & "VERIFICATION OF " & HIRE_employer_name_TIKL & "*" 
                        
                        If testing_status = True Then MsgBox list_of_TIKLs_to_delete

                        'Exit the do loop since an exact match was found
                        Exit Do
            
                    ElseIf employer_name_jobs_panel_first_word = HIRE_employer_name_first_word Then
            
                        If testing_status = True Then msgbox "Testing -- there is an exact match for employer name first word only. Will add to delete list and TIKL delete list."
            
                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & ". JOBS panel matches first word of HIRE employer name. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 
            
                        list_of_TIKLs_to_delete = list_of_TIKLs_to_delete & tikl_case_number & "-" & tikl_case_name & "-" & "VERIFICATION OF " & HIRE_employer_name_TIKL & "*" 
                        If testing_status = True Then MsgBox list_of_TIKLs_to_delete
                        
                        'Exit the do loop since an exact match was found
                        Exit Do

                    End If

                    'Ensuring that both panel_count and unea_panels_count are both numbers
                    panel_count = panel_count * 1
                    jobs_panels_count = jobs_panels_count * 1
                    
                    If panel_count = jobs_panels_count Then
                        If testing_status = True Then msgbox "Testing -- 5045 Since there were no exact employer matches, setting no_exact_JOBS_panel_matches = True"
                        'Since there were no exact employer matches, setting no_exact_JOBS_panel_matches = True
                        no_exact_JOBS_panel_matches = True
                        Exit Do
                    End If
                Loop
            End If

            'Convert string of the employer names into an array for use in the dialog
            If no_exact_JOBS_panel_matches = True Then

                'If there are 5 jobs already, it will not add another JOBS panel
                If jobs_panels_count = 5 Then
                    'Script will be unable to add another JOBS panel since there are 5 already so it will note as such and skip
                    If testing_status = True Then msgbox "Testing -- There are 5 JOBS panels already. Cannot add another JOBS panel."

                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "There does not appear to be an exactly matching JOBS panel for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & ". However, unable to add new JOBS panel for employer since there are already 5 JOBS panels. Review needed." & " Message should not be deleted."

                Else
                    'There are not 5 JOBS panels so it will check CM + 1 too before adding new JOBS panel before adding new JOBS panel
                    If testing_status = True Then MsgBox "Testing -- Need to navigate to CM + 1"
                    'PF3 back to DAIL and navigate to CASE/CURR to change the footer month and get to JOBS panel for CM
                    PF3
                    Call write_value_and_transmit("H", dail_row, 3)
                    EMReadScreen curr_panel_check, 4, 2, 55
                    If curr_panel_check <> "CURR" Then MsgBox "Testing -- not at CASE/CURR"
                    EMWriteScreen "STAT", 20, 22
                    EMWriteScreen CM_plus_1_mo, 20, 54
                    EMWriteScreen CM_plus_1_yr, 20, 57
                    Call write_value_and_transmit("JOBS", 20, 69)

                    'Open the first JOBS panel of the caregiver reference number
                    EMWriteScreen HIRE_memb_number, 20, 76
                    Call write_value_and_transmit("01", 20, 79)

                    'Read the number of JOBS panels to ensure there are not 5 already
                    EMReadScreen jobs_panels_count_CM_plus_1, 1, 2, 78

                    If jobs_panels_count_CM_plus_1 = "5" Then
                        If testing_status = True Then msgbox "Testing -- There are 5 JOBS panels already in CM + 1. Cannot add another JOBS panel."

                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "There does not appear to be an exactly matching JOBS panel for employer: " & HIRE_employer_name & " for M" & HIRE_memb_number & ". However, unable to add new JOBS panel for employer since there are already 5 JOBS panels in CM + 1. Review needed." & " Message should not be deleted."
                    Else
                        'Navigate back to CM and add JOBS as originally intended
                        'There are not 5 JOBS panels so it will check CM + 1 too before adding new JOBS panel before adding new JOBS panel
                        If testing_status = True Then MsgBox "Testing -- Need to navigate to CM + 1"
                        'PF3 back to DAIL and navigate to CASE/CURR to change the footer month and get to JOBS panel for CM
                        PF3
                        Call write_value_and_transmit("H", dail_row, 3)
                        EMReadScreen curr_panel_check, 4, 2, 55
                        If curr_panel_check <> "CURR" Then MsgBox "Testing -- not at CASE/CURR"
                        EMWriteScreen "STAT", 20, 22
                        EMWriteScreen CM_mo, 20, 54
                        EMWriteScreen CM_yr, 20, 57
                        Call write_value_and_transmit("JOBS", 20, 69)

                        'Open the first JOBS panel of the caregiver reference number
                        EMWriteScreen HIRE_memb_number, 20, 76
                        Call write_value_and_transmit("01", 20, 79)

                        Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

                        'Validation to ensure that script is able to open a new JOBS panel
                        EmReadScreen panel_count_plus_one_check, 1, 2, 73
                        panel_count_plus_one_check = panel_count_plus_one_check * 1
                        EmReadScreen panel_count_total_check, 1, 2, 78
                        panel_count_total_check = panel_count_total_check * 1

                        If panel_count_plus_one_check <> panel_count_total_check + 1 then 
                            If testing_status = True Then MsgBox "Testing -- unable to open a new JOBS panel. Will note in spreadsheet and continue"
                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "MAXIS programs are inactive. Unable to add a new JOBS panel for M" & HIRE_memb_number & ". Review needed." & " Message should not be deleted."
                        Else
                            
                            If testing_status = True Then MsgBox "Testing -- Script opened JOBS panel. Will add new panel"

                            'Reads footer month for updating the panel
                            EMReadScreen JOBS_footer_month, 2, 20, 55	
                            EMReadScreen JOBS_footer_year, 2, 20, 58	

                            'Write the date hired date from NDNH message to JOBS panel
                            Call create_MAXIS_friendly_date(date_hired, 0, 9, 35)

                            'Writes information to JOBS panel
                            EMWriteScreen "W", 5, 34
                            EMWriteScreen "4", 6, 34
                            EMWriteScreen HIRE_employer_name, 7, 42
                                
                            IF month_hired = JOBS_footer_month and year_hired = JOBS_footer_year THEN
                                'If the footer month on the JOBS panel matches the month from the HIRE message then it writes the actual hired date from the message to the panel
                                Call create_MAXIS_friendly_date(date_hired, 0, 12, 54)
                            ELSE
                                'Otherwise, write the panel footer month and date to the new panel
                                EmWriteScreen JOBS_footer_month, 12, 54
                                EMWriteScreen "01", 12, 57
                                EmWriteScreen JOBS_footer_year, 12, 60
                            END IF

                            'Puts $0 in as the received income amt and 0 worked hours
                            EMWriteScreen "0", 12, 67				
                            EMWriteScreen "0", 18, 72	
                            
                            If testing_status = True Then msgbox "Testing -- Review the JOBS panel. Any potential errors or issues before it continues?"
                            
                            'Opens FS PIC
                            Call write_value_and_transmit("X", 19, 38)
                                
                            'Write today's date to calculation since added today
                            Call create_MAXIS_friendly_date(date, 0, 5, 34)
                            
                            'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                            EMWriteScreen "1", 5, 64
                            EMWriteScreen "0", 8, 64
                            EMWriteScreen "0", 9, 66
                            If testing_status = True Then msgbox "Testing -- Review the PIC panel. Any potential errors or issues before it continues?"

                            transmit
                            EmReadScreen PIC_warning, 7, 20, 6

                            IF PIC_warning = "WARNING" then transmit 'to clear message
                            transmit 'back to JOBS panel
                            If testing_status = True Then Msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                            transmit 'to save JOBS panel

                            'Check if information is expiring and needs to be added to a future month
                            EMReadScreen expired_check, 6, 24, 17 
                            EMReadScreen data_expiration_month, 2, 24, 27
                            EMReadScreen jobs_panel_month, 2, 20, 55

                            If expired_check = "EXPIRE" THEN 
                                Do
                                    'Do loop to add JOBS panels to every month from DAIL month through CM
                                    If testing_status = True Then msgbox "Testing -- New JOBS panel is expiring so it needs to be added to CM + 1 as well"

                                    'PF3 to go to STAT/WRAP
                                    PF3

                                    'Check to make sure on STAT/WRAP
                                    EMReadScreen stat_wrap_check, 19, 2, 32
                                    If Instr(stat_wrap_check, "Wrap") = 0 Then MsgBox "Testing -- It didn't go to STAT/WRAP for some reason. Stop here!!"

                                    'Build do loop to get to expiration month so that it isn't creating a bunch of duplicate JOBS panels
                                    If data_expiration_month <> jobs_panel_month Then
                                        If testing_status = True Then msgbox "Testing -- JOBS panel expires in future month"
                                        Do
                                            Call write_value_and_transmit("Y", 16, 54)
                                            EMReadScreen stat_wrap_month_check, 2, 20, 55
                                            If stat_wrap_month_check = data_expiration_month Then
                                                'Script has reached the expiration month, it will go to next month and then exit
                                                If testing_status = True Then msgbox "Testing -- script has found matching month"
                                                PF3
                                                Call write_value_and_transmit("Y", 16, 54)
                                                Exit Do
                                            Else
                                                'Script has not yet reached the expiration month, it will PF3 back to STAT/WRAP to move to next month
                                                If testing_status = True Then msgbox "Testing -- script has not reached matching month"
                                                PF3
                                            End If
                                        Loop
                                    Else
                                        'If the expiration month and the jobs panel month are the same then it should add to next month too since it will expire at end of month
                                        Call write_value_and_transmit("Y", 16, 54)
                                    End If

                                    'Navigate to STAT/JOBS
                                    Call write_value_and_transmit("JOBS", 20, 71)

                                    EMReadScreen jobs_panel_nav_check, 8, 2, 43
                                    If InStr(jobs_panel_nav_check, "JOBS") = 0 Then MsgBox "Testing -- Stop here. Not at JOBS panel"

                                    If testing_status = True Then MsgBox "Testing -- Is it at the month after expiration? Expiration month was " & data_expiration_month

                                    'Navigate to HH member
                                    Call write_value_and_transmit(HIRE_memb_number, 20, 76)

                                    'Making sure there aren't 5 jobs already
                                    EMReadScreen five_jobs_check, 1, 2, 78
                                    
                                    If five_jobs_check = "5" Then 
                                        script_end_procedure_with_error_report("Testing -- There are 5 JOBS panels already, it will error out. Must stop here!")
                                    Else
                                        Call write_value_and_transmit("NN", 20, 79)
                                    End If
                                    
                                    EmReadScreen panel_count_plus_one_check, 1, 2, 73
                                    panel_count_plus_one_check = panel_count_plus_one_check * 1
                                    EmReadScreen panel_count_total_check, 1, 2, 78
                                    panel_count_total_check = panel_count_total_check * 1

                                    If panel_count_plus_one_check <> panel_count_total_check + 1 then script_end_procedure_with_error_report("Testing -- Unable to open a new JOBS panel. Script will stop here.")

                                    'Reads footer month for updating the panel
                                    EMReadScreen JOBS_footer_month, 2, 20, 55	
                                    EMReadScreen JOBS_footer_year, 2, 20, 58	

                                    'Write the date hired date from NDNH message to JOBS panel
                                    Call create_MAXIS_friendly_date(date_hired, 0, 9, 35)

                                    'Writes information to JOBS panel
                                    EMWriteScreen "W", 5, 34
                                    EMWriteScreen "4", 6, 34
                                    EMWriteScreen HIRE_employer_name, 7, 42
                                    
                                    'Looking at CM + 1 so won't match the message, just writes footer month to panel
                                    EmWriteScreen JOBS_footer_month, 12, 54
                                    EMWriteScreen "01", 12, 57
                                    EmWriteScreen JOBS_footer_year, 12, 60

                                    'Puts $0 in as the received income amt
                                    EMWriteScreen "0", 12, 67				
                                    'Puts 0 hours in as the worked hours
                                    EMWriteScreen "0", 18, 72		

                                    If testing_status = True Then msgbox "Testing - Does everything look good on JOBS panel before heading to PIC?"
                                    
                                    'Opens FS PIC
                                    Call write_value_and_transmit("X", 19, 38)
                                    'Writes today's date on the panel
                                    Call create_MAXIS_friendly_date(date, 0, 5, 34)

                                    'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                                    EMWriteScreen "1", 5, 64
                                    EMWriteScreen "0", 8, 64
                                    EMWriteScreen "0", 9, 66
                                    If testing_status = True Then msgbox "Testing - Does everything look good on JOBS panel before saving the PIC?"
                                    
                                    transmit
                                    EmReadScreen PIC_warning, 7, 20, 6
                                    IF PIC_warning = "WARNING" then transmit 'to clear message
                                    transmit 'back to JOBS panel
                                    If testing_status = True Then msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                                    transmit 'to save JOBS panel
                                    
                                    'Check if information is expiring and needs to be added to CM + 1
                                    EMReadScreen expired_check, 6, 24, 17 
                                    EMReadScreen data_expiration_month, 2, 24, 27
                                    EMReadScreen jobs_panel_month, 2, 20, 55 
                                    
                                    If expired_check <> "EXPIRE" THEN
                                        'If data is not expiring, then the script can exit the do loop
                                        If testing_status = True Then msgbox "Testing -- No expiration date. It will exit the do loop"
                                        Exit Do
                                    Else
                                        If testing_status = True Then msgbox "Testing -- Data is expiring. It will continue with the do loop"
                                    End If

                                Loop

                            End If

                            'Write information to CASE/NOTE
                            If testing_status = True Then MsgBox "Testing -- Script will now CASE/NOTE information. Navigate to CASE/NOTE"
                            
                            'PF4 to navigate to CASE/NOTE
                            PF4
                            
                            EMReadScreen jobs_panel_not_saved, 25, 24, 2
                            'If unable to navigate to CASE/NOTE due to not saving JOBS panel, then another transmit is needed
                            If instr(jobs_panel_not_saved, "CASE OR PERSON NOTES ARE") Then 
                                transmit
                                PF4
                            End If
                            EMReadScreen case_note_check, 4, 2, 45
                            If case_note_check <> "NOTE" then MsgBox "Testing -- not at case note stop here2"

                            'Open new CASE/NOTE
                            PF9

                            'Write information depending on whether NDNH or SDNH message
                            If InStr(dail_msg, "NDNH MEMB") Then
                                CALL write_variable_in_case_note("-NDNH Match for (M" & HIRE_memb_number & ") for " & trim(HIRE_employer_name) & "-")
                                CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
                                CALL write_variable_in_case_note("MAXIS NAME: " & NDNH_maxis_name)
                                CALL write_variable_in_case_note("NEW HIRE NAME: " & NDNH_new_hire_name)
                                CALL write_variable_in_case_note("EMPLOYER: " & HIRE_employer_name)
                                CALL write_variable_in_case_note("---")
                                CALL write_variable_in_case_note("NO CORRESPONDING JOBS PANEL EXISTED FOR EMPLOYER NOTED IN HIRE MESSAGE. STAT/JOBS PANEL ADDED FOR EMPLOYER IDENTIFIED IN HIRE DAIL MESSAGE. INFC CLEARED.")
                                CALL write_variable_in_case_note("---")
                                CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                                CALL write_variable_in_case_note("---")
                                CALL write_variable_in_case_note(worker_signature)
                            ElseIf InStr(dail_msg, "SDNH NEW JOB DETAILS") Then
                                CALL write_variable_in_case_note("-SDNH Match for (M" & HIRE_memb_number & ") for " & trim(HIRE_employer_name) & "-")
                                CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
                                CALL write_variable_in_case_note("MAXIS NAME: " & SDNH_maxis_name)
                                CALL write_variable_in_case_note("NEW HIRE NAME: " & SDNH_new_hire_name)
                                CALL write_variable_in_case_note("EMPLOYER: " & HIRE_employer_name)
                                CALL write_variable_in_case_note("---")
                                CALL write_variable_in_case_note("NO CORRESPONDING JOBS PANEL EXISTED FOR EMPLOYER NOTED IN HIRE MESSAGE. STAT/JOBS PANEL ADDED FOR EMPLOYER IDENTIFIED IN HIRE DAIL MESSAGE. HIRE MESSAGE DELETED.")
                                CALL write_variable_in_case_note("---")
                                CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                                CALL write_variable_in_case_note("---")
                                CALL write_variable_in_case_note(worker_signature)
                            Else
                                Msgbox "Testing -- something went wrong when writing the CASE/NOTE. Appears that message is neither NDNH or SDNH"
                            End If

                            If testing_status = True Then msgbox "Testing -- The script is about to save the CASE/NOTE. Stop here if in testing or production"

                            'PF3 to save the CASE/NOTE
                            PF3
                            
                            'PF3 to STAT/WRAP or JOBS
                            PF3
                            
                            EMReadScreen panel_nav_check, 4, 2, 46
                            If panel_nav_check <> "WRAP" Then
                                PF3
                                If testing_status = True Then msgbox "Testing -- The script should now be at STAT/WRAP. If it is not, then stop here."
                            End If

                            If testing_status = True Then msgbox "Testing -- No jobs panels existed. Created JOBS panel(s) through CM"
                            
                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No JOBS panels exist for household member number: " & HIRE_memb_number & " that match the HIRE message. JOBS Panel and CASE/NOTE added for employer noted in HIRE message. Message should be deleted.")
                            
                        End If
                    End If
                End If
            End if
        End If
    End If
End Function

Function nav_back_to_dail_check(testing_status)
    'Assumes script has just attempted to PF3 back to DAIL
    'Will attempt to PF3 three times before sending a msgbox
    EMReadScreen dail_panel_check, 8, 2, 46
    If InStr(dail_panel_check, "DAIL") = 0 Then 
        If testing_status = True then msgbox "Testing -- did not return to DAIL. It will PF3 again"
        PF3
        EMReadScreen dail_panel_check, 8, 2, 46
        If InStr(dail_panel_check, "DAIL") = 0 Then 
            If testing_status = True then MsgBox "Testing -- Script is still not at DAIL despite second PF3"
        End IF
    End If
End Function

'Activate message boxes
activate_msg_boxes = False

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
Call Check_for_MAXIS(False)

'Sets the county code for Hennepin County as X127
worker_county_code = "X127"

'Set footer month and year to current month and year
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

'To determine if DAIL message is in scope based on DAIL month, creating variable for date for current month, day, and year
footer_month_day_year = dateadd("d", 0, MAXIS_footer_month & "/01/20" & MAXIS_footer_year)

EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46

'Initial dialog - select whether to create a list or process a list
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 306, 260, "DAIL Unclear Information"
  GroupBox 10, 5, 290, 65, "Using the DAIL Unclear Information Script"
  Text 20, 15, 275, 50, "A BULK script that gathers and processes selected (HIRE and/or CSES) DAIL messages for the agency that fall under the Food and Nutrition Service's unclear information rules. As the DAIL messages are reviewed, the script will process DAIL messages for 6-month reporters on SNAP-only and process the DAIL messages accordingly. The data will be exported in a Microsoft Excel file type (.xlsx) and saved in the LAN. "
  Text 15, 80, 175, 10, "Type of DAIL Messages to Process:"
  CheckBox 15, 90, 55, 10, "CSES", CSES_messages
  CheckBox 15, 100, 55, 10, "HIRE", HIRE_messages
  Text 15, 115, 185, 10, "Select the X Numbers to Process (check one box only):"
  CheckBox 15, 125, 90, 10, "Process ALL X Numbers", process_all_x_numbers
  CheckBox 15, 135, 225, 10, "RESTART Process ALL X Numbers (enter restart X Number below)", restart_process_all_x_numbers
  EditBox 25, 145, 85, 15, restart_x_number
  CheckBox 15, 165, 275, 10, "RESTART Process - Enter DAIL Messages to SKIP", restart_with_skip_dail_messages
  Text 30, 175, 265, 10, "Copy exactly from spreadsheet and separate by asterisk (*) with no extra spaces"
  EditBox 25, 185, 270, 15, DAIL_messages_to_skip
  CheckBox 15, 205, 255, 10, "Process Specific X Numbers (enter X Numbers below separated by comma)", process_specific_x_numbers
  EditBox 25, 215, 270, 15, specific_x_numbers_to_process
  Text 15, 245, 60, 10, "Worker Signature:"
  EditBox 80, 240, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 240, 40, 15
    CancelButton 245, 240, 40, 15
EndDialog


DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation

        'Dialog field validation
        'Validation to ensure that at least CSES or HIRE messages checkbox is checked
        If CSES_messages = 0 AND HIRE_messages = 0 Then err_msg = err_msg & vbCr & "* Select either CSES or HIRE messages, or both. Both checkboxes cannot be left blank."
        'Validation to ensure that only one option is selected for X Numbers to process
        If process_specific_x_numbers + process_all_x_numbers + restart_process_all_x_numbers <> 1 Then err_msg = err_msg & vbCr & "* You can only select one option for processing X Numbers. Make sure only one box is checked."
        'Validation to ensure that Specific X Numbers and Restart Process All X Numbers fields are left blank if processing all X Numbers
        If process_all_x_numbers = 1 AND (trim(specific_x_numbers_to_process) <> "" OR trim(restart_x_number) <> "") Then err_msg = err_msg & vbCr & "* You selected the option to Process All X Numbers. The entry fields for Process Specific X Numbers and RESTART Process All X Numbers must be blank to proceed."
        'Validation to ensure that Process Specific X Numbers field is blank if Restart Process All X Numbers is selected
        If restart_process_all_x_numbers = 1 AND trim(specific_x_numbers_to_process) <> "" Then err_msg = err_msg & vbCr & "* You selected the option to Restart Process All X Numbers. The entry field for Process Specific X Numbers must be blank to proceed."
        If restart_process_all_x_numbers = 1 AND trim(restart_x_number) = "" Then err_msg = err_msg & vbCr & "* You selected the option to Restart Process All X Numbers. The entry field for Restart Process All X Numbers is empty. Enter the X Number that the script should restart on."
        'Validation to ensure that Restart Process All X Numbers field is blank if Process Specific X Numbers is selected
        If process_specific_x_numbers = 1 AND trim(restart_x_number) <> "" Then err_msg = err_msg & vbCr & "* You selected the option to Process Specific X Numbers. The entry field for RESTART Process All X Numbers must be empty. Clear the field to proceed."
        If restart_with_skip_dail_messages = 1 Then
			If trim(DAIL_messages_to_skip) = "" or restart_process_all_x_numbers <> 1 or trim(restart_x_number) = "" Then err_msg = err_msg & vbCr & "* You selected the option to enter DAIL messages to skip in the next run. You must enter the DAIL message to skip, as well as checking the restart process checkbox and entering the restart X number." 
		End If
		If restart_with_skip_dail_messages <> 1 and trim(DAIL_messages_to_skip) <> "" Then err_msg = err_msg & vbCr & "* You must check the DAIL messages to skip checkbox."
        'Validation to ensure that if processing specific X numbers, the list of X numbers field is not blank
        If process_specific_x_numbers = 1 AND trim(specific_x_numbers_to_process) = "" Then err_msg = err_msg & vbCr & "* You selected the option to Process Specific X Numbers. You must enter a list of X Numbers separated by a comma to proceed. The entry field is currently empty."
        'Ensures worker signature is not blank
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = "" and ButtonPressed = OK
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in	

'Handling if there are specific DAIL messages that need to be skipped
If restart_with_skip_dail_messages = 1 and trim(DAIL_messages_to_skip) <> "" Then

    'Dialog to ensure the messages to skip are formatted correctly
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 356, 250, "Messages to Skip"
    ButtonGroup ButtonPressed
        OkButton 250, 230, 50, 15
        CancelButton 300, 230, 50, 15
    GroupBox 5, 5, 345, 185, "Messages to Skip"
    Text 10, 195, 335, 25, "Ensure there are no extra spaces between messages and asterisks are used to separate the messages. Press OK to proceed if the messages are correct."
    Text 10, 15, 335, 170, DAIL_messages_to_skip
    EndDialog

    DO
        Dialog Dialog1
        cancel_confirmation

        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in	

    'If there are dail messages to skip, then set starting list_of_DAIL_messages_to_skip. Include handling to ensure that the string starts and ends in an *

    If DAIL_messages_to_skip <> "" Then 
        'Ensure first and final character in string is *
        If right(DAIL_messages_to_skip, 1) <> "*" Then DAIL_messages_to_skip = DAIL_messages_to_skip & "*"
        If left(DAIL_messages_to_skip, 1) <> "*" Then DAIL_messages_to_skip = "*" & DAIL_messages_to_skip
        list_of_DAIL_messages_to_skip = DAIL_messages_to_skip
    End If

    If activate_msg_boxes = True Then msgbox "Testing -- list_of_DAIL_messages_to_skip " & list_of_DAIL_messages_to_skip
End If

'Determining if this is a restart or not in function below when gathering the x numbers.
If restart_process_all_x_numbers = 0 then
    restart_status = False
Else 
	restart_status = True
End if 

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If process_all_x_numbers = 1 OR restart_process_all_x_numbers = 1 then
	Call create_array_of_all_active_x_numbers_in_county_with_restart(worker_array, two_digit_county_code, restart_status, restart_x_number)
Else
	x_numbers_from_dialog = split(specific_x_numbers_to_process, ", ")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x_number in x_numbers_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
End if

Call check_for_MAXIS(False)

'Set the arrays and constants so it works regardless of whether processing CSES and/or HIRE
'Create an array to track in-scope DAIL messages
DIM DAIL_message_array()

'constants for array
const dail_maxis_case_number_const      = 0
const dail_worker_const	                = 1
const dail_type_const                   = 2
const dail_month_const		            = 3
const dail_msg_const		            = 4
const full_dail_msg_const		        = 5
const dail_processing_notes_const       = 6
const dail_excel_row_const              = 7

'Create an array with PMIs to match with CASE/PERS info
Dim PMI_and_ref_nbr_array()

'Constants for the array
const ref_nbr_const           = 0
const PMI_const               = 1
const PMI_match_found_const   = 2
const hh_member_count_const   = 3


If CSES_messages = 1 Then 

    'Create an array to track case details
    DIM CSES_case_details_array()

    'constants for array
    const case_maxis_case_number_const      = 0
    const case_worker_const	                = 1
    const active_programs_const             = 2
    const pending_programs_const            = 3
    const snap_status_const                 = 4
    const snap_type_const                   = 5
    const reporting_status_const            = 6
    const sr_report_date_const              = 7 
    const recertification_date_const        = 8
    const MFIP_status_const                 = 9
    const MFIP_MFSM_review_date_const       = 10
    const MFIP_STAT_REVW_review_date_const  = 11
    const case_processing_notes_const       = 12
    const processable_based_on_case_const   = 13
    const case_excel_row_const              = 14

    'Opening the Excel file for list of DAIL messages
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True

    'Changes name of Excel sheet to DAIL Messages to capture details about in-scope DAIL messages
    ObjExcel.ActiveSheet.Name = "DAIL Messages"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "Case Number"
    objExcel.Cells(1, 2).Value = "X Number"
    objExcel.Cells(1, 3).Value = "DAIL Type"
    objExcel.Cells(1, 4).Value = "DAIL Month"
    objExcel.Cells(1, 5).Value = "DAIL Message"
    objExcel.Cells(1, 6).Value = "Full DAIL Message"
    objExcel.Cells(1, 7).Value = "Processing Notes for DAIL Message"

    FOR i = 1 to 7		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Creating second Excel sheet for compiling case details
    ObjExcel.Worksheets.Add().Name = "Case Details"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "Case Number"
    objExcel.Cells(1, 2).Value = "X Number"
    objExcel.Cells(1, 3).Value = "Active Programs"
    objExcel.Cells(1, 4).Value = "Pending Programs"
    objExcel.Cells(1, 5).Value = "SNAP Status"
    objExcel.Cells(1, 6).Value = "SNAP Type"
    objExcel.Cells(1, 7).Value = "SNAP Reporting Status"
    objExcel.Cells(1, 8).Value = "SNAP SR Report Date"
    objExcel.Cells(1, 9).Value = "SNAP Recertification Date"
    objExcel.Cells(1, 10).Value = "MFIP Status"
    objExcel.Cells(1, 11).Value = "MFIP MFSM Review Date"
    objExcel.Cells(1, 12).Value = "MFIP STAT/REVW Review Date"
    objExcel.Cells(1, 13).Value = "Processing Notes for Case"
    objExcel.Cells(1, 14).Value = "Processable based on Case Details"

    FOR i = 1 to 14		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Creates sheet to track stats for the script
    ObjExcel.Worksheets.Add().Name = "Stats"

    'Setting counters at 0
    STATS_counter = STATS_counter - 1
    not_processable_msg_count = 0
    dail_msg_deleted_count = 0
    QI_flagged_msg_count = 0

    'Enters info about runtime for the benefit of folks using the script
    objExcel.Cells(1, 1).Value = "Cases Evaluated:"
    objExcel.Cells(2, 1).Value = "Evaluated DAIL Messages:"
    objExcel.Cells(3, 1).Value = "Unprocessable DAIL Messages:"
    objExcel.Cells(4, 1).Value = "Deleted DAIL Messages:"
    objExcel.Cells(5, 1).Value = "QI Flagged Messages:"
    objExcel.Cells(6, 1).Value = "Script run time (in seconds):"
    objExcel.Cells(7, 1).Value = "Estimated time savings by using script (in minutes):"


    FOR i = 1 to 7		'formatting the cells'
        objExcel.Cells(i, 1).Font.Bold = True		'bold font'
        ObjExcel.rows(i).NumberFormat = "@" 		'formatting as text
        objExcel.columns(1).AutoFit()				'sizing the columns'
    NEXT

    'Create an array to track in-scope DAIL messages

    ReDim DAIL_message_array(7, 0)
    'Incrementor for the array
    Dail_count = 0

    'Sets variable for the Excel row to export data to Excel sheet
    dail_excel_row = 2

    ReDim CSES_case_details_array(case_excel_row_const, 0)
    'Incrementor for the array
    case_count = 0

    'Sets variable for the Excel row to export data to Excel sheet
    case_excel_row = 2

    'Reset the array 
    ReDim PMI_and_ref_nbr_array(3, 0)

    'Incrementor for the array
    member_count = 0

    deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

    For each worker in worker_array
        'Clearing out MAXIS case number so that it doesn't carry forward from previous case
        MAXIS_case_number = ""
        
        'Resetting all of the string lists
        'Creating initial string for tracking list of valid case numbers pulled from REPT/ACTV. This is used to avoid triggering a privileged case and losing connection to DAIL
        valid_case_numbers_list = "*"

        'Create list of case numbers to be used for comparison purposes as the script navigates through the DAIL
        list_of_all_case_numbers = "*"

        'Create list of DAIL messages that should be deleted. If a DAIL message matches, then it will be deleted. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc.
        list_of_DAIL_messages_to_delete = "*"

        'Create list of DAIL messages that should be skipped. If a DAIL message matches, then the script will skip past it to next DAIL row. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc. 
        If list_of_DAIL_messages_to_skip = "" then list_of_DAIL_messages_to_skip = "*"
        If activate_msg_boxes = True Then msgbox "Testing -- list_of_DAIL_messages_to_skip " & list_of_DAIL_messages_to_skip

        'Formatting the worker so there are no errors
        worker = trim(ucase(worker))

        'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason					
        back_to_self	

        Call navigate_to_MAXIS_screen("REPT", "ACTV")
        EMWriteScreen worker, 21, 13
        TRANSMIT
        EMReadScreen user_worker, 7, 21, 71
        EMReadScreen p_worker, 7, 21, 13
        IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

        IF worker_number = "X127CCL" or worker = "127CCL" THEN
            DO
                EmReadScreen worker_confirmation, 20, 3, 11 'looking for CENTURY PLAZA CLOSED
                EMWaitReady 0, 0
            LOOP UNTIL worker_confirmation = "CENTURY PLAZA CLOSED"
        END IF

        'Skips workers with no info
        EMReadScreen has_content_check, 1, 7, 8
        If has_content_check <> " " then
            'Grabbing each case number on screen
            Do
                'Set variable for next do...loop
                MAXIS_row = 7
                'Checking for the last page of cases.
                EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
                EMReadscreen number_of_pages, 4, 3, 76 'getting page number because to ensure it doesnt fail'
                number_of_pages = trim(number_of_pages)
                Do
                    EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number

                    'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
                    MAXIS_case_number = trim(MAXIS_case_number)
                    If MAXIS_case_number <> "" and instr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") <> 0 then exit do
                    valid_case_numbers_list = trim(valid_case_numbers_list & MAXIS_case_number & "*")

                    If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

                    MAXIS_row = MAXIS_row + 1
                    MAXIS_case_number = ""			'Blanking out variable
                Loop until MAXIS_row = 19
                PF8
            Loop until last_page_check = "THIS IS THE LAST PAGE"
        END IF

        'Navigates to DAIL to pull DAIL messages
        MAXIS_case_number = ""
        CALL navigate_to_MAXIS_screen("DAIL", "PICK")
        EMWriteScreen "_", 7, 39    'blank out ALL selection
        'Selects CSES DAIL Type based on dialog selection
        EMWriteScreen "X", 10, 39
        transmit

        'Enter the worker number on DAIL to pull up DAIL messages
        Call write_value_and_transmit(worker, 21, 6)
        'Transmits past not your dail message
        transmit 

        'Reads where the count of DAILs is listed. Used to verify DAIL is not empty.
        EMReadScreen number_of_dails, 1, 3, 67		

        DO
        'If this space is blank the rest of the DAIL reading is skipped
            If number_of_dails = " " Then exit do		
            'Because the script brings each new case to the top of the page, dail_row starts at 6.
            dail_row = 6	

            DO
                'Reset variables just in case they carry through
                dail_type = ""
                dail_msg = ""
                dail_month = ""
                MAXIS_case_number = ""
                actionable_dail = ""
                renewal_6_month_check = ""
                Snap_active = ""
                MFIP_active = ""
                SNAP_or_MFIP_active = ""
                Other_programs_active_or_pending = ""

                'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                EMReadScreen new_case, 8, dail_row, 63
                new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN 
                    'If there is NOT a new case number, the script will top the message
                    Call write_value_and_transmit("T", dail_row, 3)
                ELSEIF new_case = "CASE NBR" THEN
                    'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                End if

                'Resets the DAIL row since the message has now been topped
                dail_row = 6  

                'Determines the DAIL Type
                EMReadScreen dail_type, 4, dail_row, 6
                dail_type = trim(dail_type)

                'Reads the DAIL msg to determine if it is an out-of-scope message
                EMReadScreen dail_msg, 61, dail_row, 20
                dail_msg = trim(dail_msg)
                'List of out of scope messages pulled from non-actionable dails function
                If instr(dail_msg, "AMT CHILD SUPP MOD/ORD") OR _
                    instr(dail_msg, "AP OF CHILD REF NBR:") OR _
                    instr(dail_msg, "ADDRESS DIFFERS W/ CS RECORDS:") OR _
                    instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN LBUD IN THE MONTH") OR _
                    instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN SBUD IN THE MONTH") OR _
                    instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                    instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                    instr(dail_msg, "CHILD SUPP PAYMTS PD THRU THE COURT/AGENCY FOR CHILD") OR _
                    instr(dail_msg, "COMPLETE INFC PANEL") OR _
                    instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
                    instr(dail_msg, "IS COOPERATING WITH CHILD SUPPORT") OR _
                    instr(dail_msg, "IS NOT COOPERATING WITH CHILD SUPPORT") OR _
                    instr(dail_msg, "NAME DIFFERS W/ CS RECORDS:") OR _
                    instr(dail_msg, "PATERNITY ON CHILD REF NBR") OR _
                    instr(dail_msg, "REPORTED NAME CHG TO:") OR _
                    instr(dail_msg, "BENEFITS RETURNED, IF IOC HAS NEW ADDRESS") OR _
                    instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") OR _
                    instr(dail_msg, "CASE NOT AUTO-APPROVED - HRF/SR/RECERT DUE") OR _
                    instr(dail_msg, "CHANGE IN BUDGET CYCLE") OR _
                    instr(dail_msg, "COMPLETE ELIG IN FIAT") OR _
                    instr(dail_msg, "COUNTED IN LBUD AS UNEARNED INCOME") OR _
                    instr(dail_msg, "COUNTED IN SBUD AS UNEARNED INCOME") OR _
                    instr(dail_msg, "HAS EARNED INCOME IN 6 MONTH BUDGET BUT NO DCEX PANEL") OR _
                    instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") OR _
                    instr(dail_msg, "NEW ELIG RESULTS EXIST") OR _
                    instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") OR _
                    instr(dail_msg, "WARNING MESSAGES EXIST") OR _
                    instr(dail_msg, "ADDR CHG*CHK SHEL") OR _
                    instr(dail_msg, "APPLCT ID CHNGD") OR _
                    instr(dail_msg, "CASE AUTOMATICALLY DENIED") OR _
                    instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") OR _
                    instr(dail_msg, "CASE NOTE ENTERED BY") OR _
                    instr(dail_msg, "CASE NOTE TRANSFER FROM") OR _
                    instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") OR _
                    instr(dail_msg, "CASE XFER") OR _
                    instr(dail_msg, "CHANGE REPORT FORM SENT ON") OR _
                    instr(dail_msg, "DIRECT DEPOSIT STATUS") OR _
                    instr(dail_msg, "EFUNDS HAS NOTIFIED DHS THAT THIS CLIENT'S EBT CARD") OR _
                    instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") OR _
                    instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") OR _
                    instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") OR _
                    instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") OR _
                    instr(dail_msg, "NOT ACCESSED FOR 300 DAYS,SPEC NOT") OR _
                    instr(dail_msg, "PMI MERGED") OR _
                    instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") OR _
                    instr(dail_msg, "THIS CASE IS ERROR PRONE") OR _
                    instr(dail_msg, "EMPL SERV REF DATE IS > 60 DAYS; CHECK ES PROVIDER RESPONSE") OR _
                    instr(dail_msg, "LAST GRADE COMPLETED") OR _
                    instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
                    instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
                    instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") OR _
                    instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") OR _
                    instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE") OR _
                    instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") OR _
                    instr(dail_msg, "GA: STATUS IS PENDING - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "GA: STATUS IS REIN OR SUSPEND - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") or _
                    instr(dail_msg, "GRH: APPROVED VERSION EXISTS FOR JANUARY - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS") OR _
                    instr(dail_msg, "MSA RECERT DUE - NOT AUTO") or _
                    instr(dail_msg, "MSA IN PENDING STATUS - NOT AUTO") or _
                    instr(dail_msg, "APPROVED MSA VERSION EXISTS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") or _
                    instr(dail_msg, "GRH: STATUS IS REIN, PENDING OR SUSPEND - NOT AUTO") OR _
                    instr(dail_msg, "SDNH NEW JOB DETAILS FOR MEMB 00") OR _
                    instr(dail_msg, "SNAP: PENDING OR STAT EDITS EXIST") OR _
                    instr(dail_msg, "SNAP: REIN STATUS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: APPROVED VERSION ALREADY EXISTS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: AUTO-APPROVED - PREVIOUS UNAPPROVED VERSION EXISTS") OR _
                    instr(dail_msg, "SSN DIFFERS W/ CS RECORDS") OR _
                    instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                    instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED CASE WITH SANCTION") OR _
                    instr(dail_msg, "DWP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                    instr(dail_msg, "IV-D NAME DISCREPANCY") OR _
                    instr(dail_msg, "CHECK HAS BEEN APPROVED") OR _
                    instr(dail_msg, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                    instr(dail_msg, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                    instr(dail_msg, "- TRANS #") OR _
                    instr(dail_msg, "RSDI UPDATED - (REF") OR _
                    instr(dail_msg, "SSI UPDATED - (REF") OR _
                    instr(dail_msg, "SNAP ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS") then 
                        actionable_dail = False
                Else
                    actionable_dail = True
                End If

                If actionable_dail = True AND dail_type = "CSES" Then
                    'Read the MAXIS Case Number, if it is a new case number then pull case details. If it is not a new case number, then do not pull new case details.
                    
                    EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                    MAXIS_case_number = trim(MAXIS_case_number)

                    If InStr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") Then
                        'If the case is in the valid_case_numbers_list, then it can be evaluated. If it is NOT in the valid_case_numbers_list then it is likely privileged or out of county so it will be skipped

                        If Instr(list_of_all_case_numbers, "*" & MAXIS_case_number & "*") = 0 Then
                            'If the MAXIS case number is NOT in the list of all case numbers, then it is a new case number and the script will gather case details
                            'Redim the case details array and add to array
                            ReDim Preserve CSES_case_details_array(case_excel_row_const, case_count)
                            CSES_case_details_array(case_maxis_case_number_const, case_count) = MAXIS_case_number
                            CSES_case_details_array(case_worker_const, case_count) = worker
                    
                            'Since case number is not in list of all case numbers, add it to the list
                            list_of_all_case_numbers = list_of_all_case_numbers & MAXIS_case_number & "*"

                            'Navigate to CASE/CURR to pull case details 
                            Call write_value_and_transmit("H", dail_row, 3)

                            'Handling if the case is out of county
                            EmReadscreen worker_county, 4, 21, 14
                            If worker_county <> worker_county_code then
                                CSES_case_details_array(case_processing_notes_const, case_count) = "Out-of-County Case"
                                CSES_case_details_array(processable_based_on_case_const, case_count) = False
                            Else
                                'Pull case details from CASE/CURR, maintains connection to DAIL
                                Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

                                'Split list of active programs into an array to validate
                                If trim(list_active_programs) <> "" Then 
                                    split_list_active_programs = split(list_active_programs, ", ")

                                    i = 0
                                    Do
                                        If split_list_active_programs(i) = "SNAP" Then 
                                            SNAP_active = True
                                            SNAP_or_MFIP_active = True
                                        ElseIf split_list_active_programs(i) = "MFIP" Then 
                                            MFIP_active = True
                                            SNAP_or_MFIP_active = True
                                        Else
                                            'If it is a program other than SNAP, GA, and/or MFIP then we will need to skip this case
                                            other_programs_active_or_pending = other_programs_active_or_pending & split_list_active_programs(i) & ", " 
                                        End If
                                        i = i + 1
                                    Loop until i = ubound(split_list_active_programs) + 1
                                End If
                                
                                If list_pending_programs <> "" then other_programs_active_or_pending = other_programs_active_or_pending & list_pending_programs

                                If activate_msg_boxes = True Then msgbox "Testing -- SNAP_active = " & snap_active & vbcr & vbcr & "MFIP_active = " & MFIP_active & vbcr & vbcr & "GA_active = " & GA_active & vbcr & vbcr & "other_programs_active_or_pending = " & other_programs_active_or_pending  

                                'Update array with active and pending programs, and SNAP and MFIP statuses
                                CSES_case_details_array(active_programs_const, case_count) = list_active_programs
                                CSES_case_details_array(pending_programs_const, case_count) = list_pending_programs
                                CSES_case_details_array(SNAP_status_const, case_count) = trim(snap_status)
                                CSES_case_details_array(MFIP_status_const, case_count) = trim(mfip_status)

                                'Function (determine_program_and_case_status_from_CASE_CURR) sets dail_row equal to 4 so need to reset it.
                                dail_row = 6

                                If case_active = TRUE and SNAP_or_MFIP_active = True and other_programs_active_or_pending = "" Then
                                    If SNAP_active = True Then
                                        'The case is active on SNAP, will gather more details about SNAP

                                        'Ensure that we are viewing ELIG/FS for the current month, not the dail message month
                                        EMWriteScreen MAXIS_footer_month, 20, 54
                                        EMWriteScreen MAXIS_footer_year, 20, 57

                                        'Navigate to ELIG/FS from CASE/CURR to maintain tie to DAIL
                                        EMWriteScreen "ELIG", 20, 22
                                        Call write_value_and_transmit("FS  ", 20, 69)

                                        EMReadScreen no_SNAP, 10, 24, 2
                                        If no_SNAP = "NO VERSION" then						'NO SNAP version means no determination
                                            CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            CSES_case_details_array(processable_based_on_case_const, case_count) = False
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
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                            Elseif app_status = "APPROVED" then
                                                EMReadScreen vers_number, 1, status_row, 23
                                                Call write_value_and_transmit(vers_number, 18, 54)
                                                Call write_value_and_transmit("FSSM", 19, 70)

                                                EmReadscreen reporting_status, 12, 8, 31
                                                reporting_status = trim(reporting_status)

                                                'Read for UHFS
                                                EmReadscreen UHFS_status_check, 16, 4, 3
                                                If UHFS_status_check = "'UNCLE HARRY' FS" Then 
                                                    CSES_case_details_array(snap_type_const, case_count) = "UHFS"
                                                Else
                                                    CSES_case_details_array(snap_type_const, case_count) = "SNAP"
                                                End If
                                                
                                                If reporting_status = "SIX MONTH" Then
                                                    'Navigate to STAT/REVW to confirm recertification and SR report date
                                                    EMWriteScreen "STAT", 19, 22
                                                    EMWaitReady 0, 0
                                                    Call write_value_and_transmit("REVW", 19, 70)
                                                    
                                                    EMWaitReady 0, 0
                                                    EmReadscreen error_prone_check, 6, 2, 51

                                                    If InStr(error_prone_check, "ERRR") Then
                                                        transmit
                                                        EMWaitReady 0, 0
                                                    End If

                                                    'Pause here as it sometimes errors
                                                    EMWaitReady 0, 0
                                                    'Open the FS screen
                                                    EMWriteScreen "X", 5, 58
                                                    EMWaitReady 0, 0
                                                    Transmit
                                                    EMWaitReady 0, 0

                                                    EMReadScreen food_support_reports_check, 20, 5, 30
                                                    If food_support_reports_check <> "Food Support Reports" Then 
                                                        'Pause here as it sometimes errors
                                                        EMWaitReady 0, 0
                                                        'Open the FS screen
                                                        EMWriteScreen "X", 5, 58
                                                        EMWaitReady 0, 0
                                                        Transmit
                                                        EMWaitReady 0, 0
                                                        EMReadScreen food_support_reports_check, 20, 5, 30
                                                        If food_support_reports_check <> "Food Support Reports" Then MsgBox "Testing -- FS Screen attempt 2 did not work. Try rerunning script again."
                                                    End If

                                                    EmReadscreen sr_report_date, 8, 9, 26
                                                    EmReadscreen recertification_date, 8, 9, 64

                                                    'Add handling for missing SR report date or recertification
                                                    'Adds slashes to dates then converts to datedate from string to date
                                                    If sr_report_date = "__ 01 __" Then
                                                        sr_report_date = "SR Report Date is Missing"
                                                    Else
                                                        sr_report_date = replace(sr_report_date, " ", "/")
                                                        sr_report_date = DateAdd("m", 0, sr_report_date)
                                                    End If

                                                    If recertification_date = "__ 01 __" Then
                                                        recertification_date = "Recertification Date is Missing"
                                                    Else
                                                        recertification_date = replace(recertification_date, " ", "/")
                                                        recertification_date = DateAdd("m", 0, recertification_date)
                                                    End If
                            
                                                    If sr_report_date <> "SR Report Date is Missing" and recertification_date <> "Recertification Date is Missing" Then 
                                                        renewal_6_month_difference = DateDiff("M", sr_report_date, recertification_date)

                                                        If renewal_6_month_difference = "6" or renewal_6_month_difference = "-6" then 
                                                            renewal_6_month_check = True
                                                        Else 
                                                            renewal_6_month_check = False
                                                            CSES_case_details_array(case_processing_notes_const, case_count) = "SR Report Date and Recertification are not 6 months apart"
                                                        End if

                                                        If DateDiff("m", footer_month_day_year, sr_report_date) < 0 and DateDiff("m", footer_month_day_year, recertification_date) < 0 Then
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Both SR Report and Recert Dates are before CM and has not been updated correctly 1543"
                                                            If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; SNAP Review Dates are prior to current month. Case should be reviewed."
                                                            Else
                                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "SNAP Review Dates are prior to current month. Case should be reviewed."
                                                            End If
                                                            CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                                        End If
                                                    Else
                                                        renewal_6_month_check = False
                                                        CSES_case_details_array(case_processing_notes_const, case_count) = "SR Report Date and/or Recertification Date is missing"
                                                    End If
                                                    
                                                    'Close the FS screen
                                                    transmit
                                                Else
                                                    sr_report_date = "N/A"
                                                    recertification_date = "N/A"
                                                End If
                                            End If
                                            
                                            'Update the array with new case details
                                            CSES_case_details_array(reporting_status_const, case_count) = reporting_status
                                            CSES_case_details_array(recertification_date_const, case_count) = trim(recertification_date)
                                            CSES_case_details_array(sr_report_date_const, case_count) = trim(sr_report_date)

                                        End If
                                    End If

                                    If MFIP_active = True Then
                                        'Navigate to MFSM panel to confirm review date
                                        'Navigate to STAT/REVW to confirm review date

                                        'To ensure starting from DAIL, PF3 to get back to DAIL then navigate back to CASE/CURR
                                        'Back to DAIL
                                        PF3
                                        'Navigate back to CASE/CURR
                                        Call write_value_and_transmit("H", dail_row, 3)
                                        'Update the footer month/year and then navigate to ELIG/GA

                                        'Ensure that we are viewing ELIG/FS for the current month, not the dail message month
                                        EMWriteScreen MAXIS_footer_month, 20, 54
                                        EMWriteScreen MAXIS_footer_year, 20, 57
                                        
                                        'Navigate to ELIG/GA from CASE/CURR
                                        EMWriteScreen "ELIG", 20, 22
                                        Call write_value_and_transmit("MFIP", 20, 69)

                                        EMReadScreen no_MFIP, 10, 24, 2
                                        If no_MFIP = "NO VERSION" then						'NO GA version means no determination
                                            If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; No version of MFIP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            Else
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "No version of MFIP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            End If
                                            CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                        Else
                                            EMWriteScreen "99", 20, 79
                                            transmit
                                            'This brings up the MFIP versions of eligibility results to search for approved versions
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
                                                If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                    CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; No approved eligibility results for MFIP exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                Else
                                                    CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "No approved eligibility results for MFIP exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                End If
                                                CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                            Elseif app_status = "APPROVED" then
                                                'View approved eligibility results
                                                EMReadScreen vers_number, 1, status_row, 23
                                                Call write_value_and_transmit(vers_number, 18, 54)
                                                'Navigate to MFSM panel to read the eligibility review date
                                                Call write_value_and_transmit("MFSM", 20, 71)
                                                EmReadScreen MFSM_panel_check, 4, 3, 47
                                                If MFSM_panel_check <> "MFSM" Then msgbox "Testing -- 1623 Error unable to reach MFSM"
                                                
                                                'Read eligibility review date from MFSM panel
                                                EMReadScreen MFIP_MFSM_review_date, 8, 11, 31
                                                CSES_case_details_array(MFIP_MFSM_review_date_const, case_count) = trim(MFIP_MFSM_review_date)
                                                
                                                'Navigate to STAT/REVW to confirm review date there
                                                EMWriteScreen "STAT", 20, 13
                                                Call write_value_and_transmit("REVW", 20, 71)
                                                
                                                EmReadScreen REVW_panel_check, 4, 2, 46
                                                ' If REVW_panel_check <> "REVW" Then msgbox "Testing -- 1634 Error unable to reach STAT/REVW"
                                                
                                                'Open the CASH/GRH window
                                                Call write_value_and_transmit("X", 5, 35)
                                                'Read eligibility review date 
                                                EMReadScreen MFIP_STAT_REVW_review_date, 8, 9, 64
                                                'If the review date is blank, then the case should be flagged and skipped for processing
                                                If Instr(MFIP_STAT_REVW_review_date, "_") Then
                                                    If activate_msg_boxes = True Then msgbox "Delete after Testing -- error, review date on STAT/REVW for MFIP is empty 1642"
                                                    CSES_case_details_array(MFIP_MFSM_review_date_const, case_count) = trim(MFIP_MFSM_review_date)
                                                    If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                        CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; MFIP - ER Report Date is blank on STAT/REVW"
                                                    Else
                                                        CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "MFIP - ER Report Date is blank on STAT/REVW"
                                                    End If
                                                    CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                                Else
                                                    'ER Report date is filled out so convert to MM/DD/YY
                                                    MFIP_STAT_REVW_review_date = replace(MFIP_STAT_REVW_review_date, " ", "/")

                                                    'Update the array
                                                    CSES_case_details_array(MFIP_STAT_REVW_review_date_const, case_count) = trim(MFIP_STAT_REVW_review_date)

                                                    'Compare the review date from MFSM and from STAT/REVW to identify any discrepancies
                                                    MFIP_MFSM_review_date = dateadd("d", 0, MFIP_MFSM_review_date)      'Convert to date
                                                    MFIP_STAT_REVW_review_date = dateadd("d", 0, MFIP_STAT_REVW_review_date)      'Convert to date
                                                    If MFIP_STAT_REVW_review_date <> MFIP_MFSM_review_date Then
                                                        If activate_msg_boxes = True Then msgbox "Testing -- STAT/REVW does not match MFSM 1662"
                                                        If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                            CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; Eligibility Review Date on MFSM does not match ER Report Date on STAT/REVW"
                                                        Else
                                                            CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "Eligibility Review Date on MFSM does not match ER Report Date on STAT/REVW"
                                                        End If
                                                        CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                                    End If
                                                    If MFIP_STAT_REVW_review_date = MFIP_MFSM_review_date Then
                                                        If DateDiff("m", footer_month_day_year, MFIP_STAT_REVW_review_date) < 0 Then
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review date is before CM and has not been updated correctly 1670"
                                                            If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; MFIP Review Date is prior to current month. Case should be reviewed."
                                                            Else
                                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "MFIP Review Date is prior to current month. Case should be reviewed."
                                                            End If
                                                            CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                                        End If
                                                    End If

                                                End If
                                                'Close the CASH screen
                                                transmit
                                            End If
                                        End If
                                    End If
                                Else
                                    'Case is not processable. Write information to array accordingly
                                    CSES_case_details_array(snap_type_const, case_count) = "N/A"
                                    CSES_case_details_array(reporting_status_const, case_count) = "N/A"
                                    CSES_case_details_array(sr_report_date_const, case_count) = "N/A"
                                    CSES_case_details_array(recertification_date_const, case_count) = "N/A"
                                    CSES_case_details_array(MFIP_MFSM_review_date_const, case_count) = "N/A"
                                    CSES_case_details_array(MFIP_STAT_REVW_review_date_const, case_count) = "N/A"
                                    CSES_case_details_array(case_processing_notes_const, case_count) = "Not processable"
                                    CSES_case_details_array(processable_based_on_case_const, case_count) = False
                                End If
                            End If    

                            'Only need to check if case is processable if it has not already been determined to be not processable
                            If CSES_case_details_array(processable_based_on_case_const, case_count) <> False or trim(CSES_case_details_array(processable_based_on_case_const, case_count)) = "" Then
                                'Handling for SNAP, check if SNAP is active, if it is then verify it meets criteria
                                If SNAP_active = True Then
                                    If CSES_case_details_array(snap_type_const, case_count) = "SNAP" Then
                                        If CSES_case_details_array(snap_status_const, case_count) <> "ACTIVE" OR CSES_case_details_array(reporting_status_const, case_count) <> "SIX MONTH" OR renewal_6_month_check <> True then
                                            If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; SNAP Not Processable"
                                            Else
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "SNAP Not Processable"
                                            End If
                                        End If            
                                    ElseIf CSES_case_details_array(snap_type_const, case_count) = "UHFS" Then
                                        If CSES_case_details_array(snap_status_const, case_count) <> "ACTIVE" OR CSES_case_details_array(reporting_status_const, case_count) <> "SIX MONTH" then
                                            If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; UHFS Not Processable"
                                            Else
                                                CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "UHFS Not Processable"
                                            End If
                                        End If
                                    Else
                                        msgbox "Testing -- 1705 missing some handling here. Shouldn't be hitting these, right?"
                                        If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                            CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; SNAP or UHFS Not Processable"
                                        Else
                                            CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "SNAP or UHFS Not Processable"
                                        End If
                                    End If
                                End If

                                If MFIP_active = True Then
                                    If CSES_case_details_array(MFIP_status_const, case_count) <> "ACTIVE" then
                                        If CSES_case_details_array(case_processing_notes_const, case_count) <> "" Then 
                                            CSES_case_details_array(CSES_case_details_arraycase_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "; MFIP Not Processable"
                                        Else
                                            CSES_case_details_array(case_processing_notes_const, case_count) = CSES_case_details_array(case_processing_notes_const, case_count) & "MFIP Not Processable"
                                        End If
                                    End If
                                End If
                            End If

                            If trim(CSES_case_details_array(case_processing_notes_const, case_count)) <> "" Then
                                CSES_case_details_array(processable_based_on_case_const, case_count) = False
                            ElseIf trim(CSES_case_details_array(case_processing_notes_const, case_count)) = "" Then
                                CSES_case_details_array(processable_based_on_case_const, case_count) = True
                            End If
                            
                            'Activate the case details sheet
                            objExcel.Worksheets("Case Details").Activate

                            'Update the Case Details sheet with case data
                            objExcel.Cells(case_excel_row, 1).Value = CSES_case_details_array(case_maxis_case_number_const, case_count)
                            objExcel.Cells(case_excel_row, 2).Value = CSES_case_details_array(case_worker_const, case_count)
                            objExcel.Cells(case_excel_row, 3).Value = CSES_case_details_array(active_programs_const, case_count)
                            objExcel.Cells(case_excel_row, 4).Value = CSES_case_details_array(pending_programs_const, case_count)
                            objExcel.Cells(case_excel_row, 5).Value = CSES_case_details_array(snap_status_const, case_count)
                            objExcel.Cells(case_excel_row, 6).Value = CSES_case_details_array(snap_type_const, case_count)
                            objExcel.Cells(case_excel_row, 7).Value = CSES_case_details_array(reporting_status_const, case_count)
                            objExcel.Cells(case_excel_row, 8).Value = CSES_case_details_array(sr_report_date_const, case_count)
                            objExcel.Cells(case_excel_row, 9).Value = CSES_case_details_array(recertification_date_const, case_count)
                            objExcel.Cells(case_excel_row, 10).Value = CSES_case_details_array(MFIP_status_const, case_count)
                            objExcel.Cells(case_excel_row, 11).Value = CSES_case_details_array(MFIP_MFSM_review_date_const, case_count)
                            objExcel.Cells(case_excel_row, 12).Value = CSES_case_details_array(MFIP_STAT_REVW_review_date_const, case_count)
                            objExcel.Cells(case_excel_row, 13).Value = CSES_case_details_array(case_processing_notes_const, case_count)
                            objExcel.Cells(case_excel_row, 14).Value = CSES_case_details_array(processable_based_on_case_const, case_count)

                            If CSES_case_details_array(processable_based_on_case_const, case_count) = True and activate_msg_boxes = True Then msgbox "Delete after testing -- Script found case that is in-scope, double-check spreadsheet"
                            case_excel_row = case_excel_row + 1

                            EmReadScreen case_curr_check, 4, 2, 55
                            If case_curr_check = "CURR" Then
                                EMWriteScreen MAXIS_footer_month, 20, 54
                                EMWriteScreen MAXIS_footer_year, 20, 57
                                'PF3 back to DAIL
                                PF3 
                            Else
                                'Return to DAIL by PF3
                                PF3

                                'Reset the footer month/year to CM through CASE/CURR
                                Call write_value_and_transmit("H", dail_row, 3)
                                EMWriteScreen MAXIS_footer_month, 20, 54
                                EMWriteScreen MAXIS_footer_year, 20, 57
                                PF3
                            End If
                            
                            'Increment the case_count for updating the array
                            case_count = case_count + 1
                            'Subtract one from dail_row so that the dail_row restarts evaluation of cases now with case details
                            dail_row = dail_row - 1
                        
                        Else
                            'If the MAXIS case number IS in the list of all case numbers, then it is not a new case number and no case details need to be gathered. It can work off the already collected case details.

                            'Before determining whether the DAIL is processable, script determines if it has encountered this DAIL message previously. Based on determination, it then processes (deletes) the dail, skips it, or makes processable determination

                            'Resetting the full_dail_msg to ensure it is not carrying forward to subsequent loops
                            full_dail_msg = ""

                            'Script opens the entire DAIL message to evaluate if it is a new message or not
                            Call write_value_and_transmit("X", dail_row, 3)

                            'Handling for reading full dail message depends on message type

                            If dail_type = "CSES" Then

                                'Check if the full message is displayed
                                EMReadScreen full_message_check, 36, 24, 2
                                If InStr(full_message_check, "THE ENTIRE MESSAGE TEXT") Then
                                    EMReadScreen dail_msg, 61, dail_row, 20
                                    dail_msg = trim(dail_msg)
                                    full_dail_msg = dail_msg
                                    EMWriteScreen " ", dail_row, 3
                                Else
                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                    EMReadScreen full_dail_msg_line_1, 60, 9, 5
                                    EMReadScreen full_dail_msg_line_2, 60, 10, 5
                                    EMReadScreen full_dail_msg_line_3, 60, 11, 5
                                    full_dail_msg_line_3 = trim(full_dail_msg_line_3)
                                    EMReadScreen full_dail_msg_line_4, 60, 12, 5
                                    full_dail_msg_line_4 = trim(full_dail_msg_line_4)

                                    If trim(full_dail_msg_line_2) = "" Then 
                                        full_dail_msg_line_1 = trim(full_dail_msg_line_1)
                                    End If

                                    full_dail_msg = trim(full_dail_msg_line_1 & full_dail_msg_line_2 & full_dail_msg_line_3 & full_dail_msg_line_4)

                                    'Transmit back to dail
                                    transmit

                                End If
                            End If

                            'The script has the full DAIL message and can compare against delete and skip lists to determine if it is a new message

                            If Instr(list_of_DAIL_messages_to_delete, "*" & full_dail_msg & "*") Then
                                'If the full dail message is within the list of dail messages to delete then the message should be deleted

                                'Resetting variables so they do not carry forward
                                last_dail_check = ""
                                other_worker_error = ""
                                total_dail_msg_count_before = ""
                                total_dail_msg_count_after = ""
                                all_done = ""
                                final_dail_error = ""
                                
                                'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
                                EMReadScreen last_dail_check, 12, 3, 67
                                last_dail_check = trim(last_dail_check)

                                'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
                                last_dail_check = split(last_dail_check, " ")

                                If last_dail_check(0) = last_dail_check(2) then 
                                    'The script is about to delete the LAST message in the DAIL so script will exit do loop after deletion, also works if it is about to delete the ONLY message in the DAIL
                                    all_done = true
                                End If

                                'Delete after testing new functionality
                                ' activate_msg_boxes = True
                                If activate_msg_boxes = True Then MsgBox "It is about to delete the message. Confirm before proceeding."
                                'Delete the message
                                Call write_value_and_transmit("D", dail_row, 3)
                                ' activate_msg_boxes = False

                                'Handling for deleting message under someone else's x number
                                EMReadScreen other_worker_error, 25, 24, 2
                                other_worker_error = trim(other_worker_error)

                                If other_worker_error = "ALL MESSAGES WERE DELETED" Then
                                    'Script deleted the final message in the DAIL
                                    dail_row = dail_row - 1
                                    dail_msg_deleted_count = dail_msg_deleted_count + 1
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    'Exit do loop as all messages are deleted
                                    all_done = true

                                ElseIf other_worker_error = "" Then
                                    'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion

                                    'Handling to check if message actually deleted
                                    total_dail_msg_count_before = last_dail_check(2) * 1
                                    EMReadScreen total_dail_msg_count_after, 12, 3, 67

                                    total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                    total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                    If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                        'The total DAILs decreased by 1, message deleted successfully
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    Else
                                        'The total DAILs did not decrease by 1, something went wrong
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 881.")
                                    End If

                                ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
                                    
                                    'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
                                    transmit

                                    'Reads the total number of DAILS after deleting to determine if it decreased by 1
                                    EMReadScreen total_dail_msg_count_after, 12, 3, 67

                                    'Checks if final DAIL message deleted
                                    EMReadScreen final_dail_error, 25, 24, 2

                                    If final_dail_error = "ALL MESSAGES WERE DELETED" Then
                                        'All DAIL messages deleted so indicates deletion a success
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        'No more DAIL messages so exit do loop
                                        all_done = True
                                    ElseIf trim(final_dail_error) = "" Then
                                        'Handling to check if message actually deleted
                                        total_dail_msg_count_before = last_dail_check(2) * 1

                                        total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                        total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                        If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                            'The total DAILs decreased by 1, message deleted successfully
                                            dail_row = dail_row - 1
                                            dail_msg_deleted_count = dail_msg_deleted_count + 1
                                            objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        Else
                                            objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                            script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 915.")
                                        End If

                                    Else
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 920.")
                                    End if
                                    
                                Else
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 925.")
                                End If

                                If activate_msg_boxes = True Then MsgBox "The message has been deleted. Did anything go wrong? If so, stop here!"
                            ElseIf Instr(list_of_DAIL_messages_to_skip, "*" & full_dail_msg & "*") Then
                                'If the full message is on the list of dail messages to skip then the message should be skipped

                                If Instr(DAIL_messages_to_skip, "*" & full_dail_msg & "*") and activate_msg_boxes = True then msgbox "It hit a message and skipped it that was on list of skips -> " & full_dail_msg

                            ElseIf Instr(list_of_DAIL_messages_to_delete, "*" & full_dail_msg & "*") = 0 AND Instr(list_of_DAIL_messages_to_skip, "*" & full_dail_msg & "*") = 0 Then
                                'If the full dail message is NOT in the list of dail messages to delete AND the full dail messages is NOT in the list of skip messages then it SHOULD be a new dail message and therefore it needs to be evaluated

                                'Gather details on DAIL message, should capture DAIL details in spreadsheet even if ultimately not actionable
                            
                                'Reset the array
                                ReDim Preserve DAIL_message_array(DAIL_excel_row_const, dail_count)
                                DAIL_message_array(dail_maxis_case_number_const, DAIL_count) = MAXIS_case_number
                                DAIL_message_array(dail_worker_const, DAIL_count) = worker

                                'Use for next loop to match the individual DAIL message to the corresponding array item of matching Case Details
                                for each_case = 0 to UBound(CSES_case_details_array, 2)
                                    'Iterate through each of the cases 
                                    If DAIL_message_array(dail_maxis_case_number_const, dail_count) = CSES_case_details_array(case_maxis_case_number_const, each_case) Then
                                        'As the for to loop iterates through each case details array, if the dail maxis case number for the dail message array matches the maxis case number for the case details array then it can pull the case details from the array  
                                        
                                        'Clearing out process_dail_message
                                        process_dail_message = ""

                                        'Read dail message details
                                        EMReadScreen dail_type, 4, dail_row, 6
                                        dail_type = trim(dail_type)

                                        EMReadScreen dail_month, 8, dail_row, 11
                                        dail_month = trim(dail_month)

                                        EMReadScreen dail_msg, 61, dail_row, 20
                                        dail_msg = trim(dail_msg)

                                        'Update the DAIL message array with details
                                        DAIL_message_array(dail_type_const, dail_count) = dail_type
                                        DAIL_message_array(dail_month_const, dail_count) = dail_month
                                        DAIL_message_array(dail_msg_const, dail_count) = dail_msg
                                        DAIL_message_array(full_dail_msg_const, dail_count) = full_dail_msg

                                        'Activate the DAIL Messages sheet
                                        objExcel.Worksheets("DAIL Messages").Activate

                                        'Write dail details to the Excel sheet
                                        objExcel.Cells(dail_excel_row, 1).Value = DAIL_message_array(dail_maxis_case_number_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 2).Value = DAIL_message_array(dail_worker_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 3).Value = DAIL_message_array(dail_type_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 4).Value = DAIL_message_array(dail_month_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 5).Value = DAIL_message_array(dail_msg_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 6).Value = DAIL_message_array(full_dail_msg_const, dail_count)

                                        If CSES_case_details_array(processable_based_on_case_const, each_case) = False Then
                                            If Instr(CSES_case_details_array(case_processing_notes_const, each_case), "SR Report Date and Recertification are not 6 months apart") OR _
                                                Instr(CSES_case_details_array(case_processing_notes_const, each_case), "SR Report Date and/or Recertification Date is missing") OR _
                                                Instr(CSES_case_details_array(case_processing_notes_const, each_case), "SNAP Review Dates are prior to current month. Case should be reviewed") OR _
                                                Instr(CSES_case_details_array(case_processing_notes_const, each_case), "MFIP - ER Report Date is blank on STAT/REVW") OR _
                                                Instr(CSES_case_details_array(case_processing_notes_const, each_case), "Eligibility Review Date on MFSM does not match ER Report Date on STAT/REVW") OR _ 
                                                Instr(CSES_case_details_array(case_processing_notes_const, each_case), "MFIP Review Date is prior to current month. Case should be reviewed") Then 
                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "QI review needed." & CSES_case_details_array(case_processing_notes_const, each_case)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                            Else
                                                DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable based on Case Details: " & CSES_case_details_array(case_processing_notes_const, each_case)
                                                not_processable_msg_count = not_processable_msg_count + 1
                                            End If

                                            'The dail message should not be processed due to case details
                                            process_dail_message = False

                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                            'Activate the DAIL Messages sheet
                                            objExcel.Worksheets("DAIL Messages").Activate

                                            'Update the Excel sheet
                                            objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                        
                                        ElseIf CSES_case_details_array(processable_based_on_case_const, each_case) = True Then     

                                            'Handling for SNAP/UHFS to check if review or recert is CM + 1. If so, checks if DAIL month is CM + 1 too. If that's the case, it will skip processing the message.
                                            If CSES_case_details_array(snap_type_const, each_case) = "SNAP" or CSES_case_details_array(snap_type_const, each_case) = "UHFS" Then
                                                'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                                If DateAdd("m", 0, CSES_case_details_array(recertification_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, CSES_case_details_array(sr_report_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) Then
                                                    If activate_msg_boxes = True Then Msgbox "The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"

                                                    If dail_type = "CSES" Then
                                                        
                                                        If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 1, footer_month_day_year) Then

                                                            DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal for SNAP/UHFS. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                            objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                            not_processable_msg_count = not_processable_msg_count + 1

                                                            'The dail message cannot be processed due to timing of recertification or SR report date
                                                            process_dail_message = False

                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        Else

                                                            'Process the CSES message here
                                                            process_dail_message = True

                                                        End If
                                                    End If

                                                Else
                                                    'If neither the recertification or SR report date is next month then we assume the dail message can be processed since processable based on case details is True. So set the process_dail_message to True to gather more information about the dail message
                                                    process_dail_message = True
                                                    
                                                End If
                                            End If

                                            'Handling for MFIP to check if review or recert is CM + 1. If so, checks if DAIL month is CM + 1 too. If that's the case, it will skip processing the message.
                                            If CSES_case_details_array(MFIP_status_const, each_case) = "ACTIVE" Then
                                                'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                                'Subtract 6 months from ER Report Date to get review date
                                                ER_report_minus_6_months = DateAdd("m", -6, CSES_case_details_array(MFIP_STAT_REVW_review_date_const, each_case))

                                                If DateAdd("m", 0, CSES_case_details_array(MFIP_STAT_REVW_review_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, ER_report_minus_6_months) = DateAdd("m", 1, footer_month_day_year) Then
                                                    If activate_msg_boxes = True Then Msgbox "The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"

                                                    If dail_type = "CSES" Then
                                                        
                                                        If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 1, footer_month_day_year) Then

                                                            If trim(DAIL_message_array(dail_processing_notes_const, dail_count)) = "" then
                                                                DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal for MFIP. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                            Else
                                                                DAIL_message_array(dail_processing_notes_const, dail_count) = DAIL_message_array(dail_processing_notes_const, dail_count) & "; Not Processable due to DAIL Month & Recert/Renewal for MFIP. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                            End If
                                                            
                                                            objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                            not_processable_msg_count = not_processable_msg_count + 1

                                                            'The dail message cannot be processed due to timing of recertification or SR report date
                                                            process_dail_message = False
                                                            'Add to skip list
                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        Else
                                                            'DAIL message can be processed
                                                            process_dail_message = True
                                                        End If
                                                    End If

                                                Else
                                                    'If neither the recertification or SR report date is next month then we assume the dail message can be processed since processable based on case details is True. So set the process_dail_message to True to gather more information about the dail message
                                                    process_dail_message = True
                                                    
                                                End If
                                            End If

                                            'Process the CSES dail message
                                            If process_dail_message = True and dail_type = "CSES" Then

                                                If InStr(dail_msg, "DISB CS (TYPE 36) OF") Then

                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                    EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                    EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                    EMReadScreen check_full_dail_msg_line_4, 60, 12, 5

                                                    If trim(check_full_dail_msg_line_2) = "" Then 
                                                        check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                    End If

                                                    check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- check_full_dail_msg <> full_dail_msg. STOP HERE something went wrong"
                                                    End if

                                                    ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                    'Identify where 'PMI(S):' text is so that script can account for Type 36 and replaced Type 36 is
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "PMI(S):", row, col
                                                    EMReadScreen PMIs_line_one, 65 - (col + 8), row, col + 8
                                                    EMReadScreen PMIs_line_two, 60, 11, 5
                                                    EMReadScreen PMIs_line_three, 60, 12, 5
                                                    
                                                    'Combine the PMIs into one string
                                                    full_PMIs = replace(PMIs_line_one & PMIs_line_two & PMIs_line_three, " ", "")
                                                    'Splits the PMIs into an array
                                                    PMIs_array = Split(full_PMIs, ",")

                                                    'Reset the array 
                                                    ReDim PMI_and_ref_nbr_array(3, 0)

                                                    'Using list of PMIs in PMIs_array to update the PMI number in PMI_and_ref_nbr_array 
                                                    for each_PMI = 0 to UBound(PMIs_array, 1)
                                                        ReDim Preserve PMI_and_ref_nbr_array(hh_member_count_const, each_PMI)
                                                        PMI_and_ref_nbr_array(PMI_const, each_PMI) = PMIs_array(each_PMI)
                                                    Next 

                                                    'Transmit back to DAIL
                                                    transmit

                                                    ' Navigate to CASE/PERS to match PMIs and Ref Nbrs for checking UNEA panel
                                                    Call write_value_and_transmit("H", dail_row, 3)

                                                    EMWriteScreen "PERS", 20, 69
                                                    Transmit

                                                    ' Iterate through CASE/PERS to match up PMI with Ref Nbr

                                                    'the first member number starts at row 10
                                                    pers_row = 10                  

                                                    Do
                                                        'Reset reference number and PMI number so they don't carry through loop
                                                        ref_number_pers_panel = ""
                                                        pmi_number_pers_panel = ""

                                                        'reading the member number
                                                        EMReadScreen ref_number_pers_panel, 2, pers_row, 3
                                                        ref_number_pers_panel = trim(ref_number_pers_panel)

                                                        'Reading the PMI number
                                                        EMReadScreen pmi_number_pers_panel, 8, pers_row, 34  
                                                        pmi_number_pers_panel = trim(pmi_number_pers_panel)

                                                        for each_PMI = 0 to UBound(PMI_and_ref_nbr_array, 2)

                                                            If pmi_number_pers_panel = PMI_and_ref_nbr_array(PMI_const, each_PMI) Then
                                                                PMI_and_ref_nbr_array(ref_nbr_const, each_PMI) = ref_number_pers_panel
                                                                PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI) = True
                                                            End If
                                                        Next
                                                        
                                                        'go to the next member number - which is 3 rows down
                                                        pers_row = pers_row + 3

                                                        'if it reaches 19 - this is further down from the last member
                                                        If pers_row = 19 Then                       
                                                            'go to the next page and reset to line 10
                                                            PF8         
                                                            EMReadScreen last_page_check, 21, 24, 2                          
                                                            If last_page_check = "THIS IS THE LAST PAGE" Then Exit Do   
                                                            pers_row = 10
                                                        End If

                                                        EMReadScreen ref_number_pers_panel, 2, pers_row, 3

                                                    Loop until ref_number_pers_panel = "  "      
                                                    
                                                    'If there are PMIs listed on the DAIL message that do not align, then that DAIL message needs to be flagged for QI
                                                    for each_individual = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                        If PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) <> True Then
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " not found on case.")
                                                        ElseIf PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) = True Then
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " found on case (M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ").")
                                                        Else
                                                            MsgBox "Testing -- Script unable to determine if all or only some PMIs matched. STOP HERE. Something went wrong."
                                                        End If
                                                    Next

                                                    'Only check UNEA panels if ALL PMIs are matched for DAIL message and for case. There are no PMIs that did not match within the array.
                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case") = 0 Then
                                                        'If all PMIs are found on the case, then the script will navigate directly to STAT/UNEA from CASE/PERS to verify that UNEA panels exist for CS Type 36 for each identified PMI/reference number

                                                        'Update the processing notes to indicate that all PMIs found on the case rather than listing out on by one
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = "All PMIs found on case. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        EMWriteScreen "STAT", 19, 22
                                                        Call write_value_and_transmit("UNEA", 19, 69)

                                                        EmReadScreen no_unea_panels_exist, 34, 24, 2

                                                        If trim(no_unea_panels_exist) = "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                            'If no UNEA panels exist for the case, then the case needs to be flagged for QI
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = " No UNEA panels exist for any member on the case." & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        ElseIf trim(no_unea_panels_exist) <> "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                            'There are at least some UNEA panels. Iterate through all of the PMI/reference numbers to ensure there are corresponding UNEA panels for the DISB Type
                                                            for each_individual = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                                'Navigate to first UNEA panel for member to determine if any exist
                                                                EMWriteScreen PMI_and_ref_nbr_array(ref_nbr_const, each_individual), 20, 76
                                                                Call write_value_and_transmit("01", 20, 79)

                                                                'Check if no UNEA panel exists
                                                                EmReadScreen unea_panel_check, 25, 24, 2

                                                                If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                                    'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ".")
                                                                Else
                                                                    'Read the UNEA type
                                                                    EMReadScreen unea_type, 2, 5, 37
                                                                    If unea_type = "36" Then
                                                                        'If it is a type 36 panel then it does not need to read the other panels
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                    Else
                                                                        'Check how many panels exist for the HH member
                                                                        EMReadScreen unea_panels_count, 1, 2, 78

                                                                        If unea_panels_count = "1" Then
                                                                            'If there is only one UNEA panel and it is not a Type 36 then it will update processing notes
                                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                            
                                                                        ElseIf unea_panels_count <> "1" Then
                                                                            'If there are more than just a single UNEA panel, loop through them all to check for Type 36
                                                                            'Set incrementor for do loop
                                                                            panel_count = 1

                                                                            Do
                                                                                panel_count = panel_count + 1
                                                                                EMWriteScreen PMI_and_ref_nbr_array(ref_nbr_const, each_individual), 20, 76
                                                                                Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                                'Read the UNEA type
                                                                                EMReadScreen unea_type, 2, 5, 37
                                                                                If unea_type = "36" Then
                                                                                    'If it is a type 36 panel then it does not need to read the other panels
                                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                    Exit Do
                                                                                End if

                                                                                'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                                panel_count = panel_count * 1
                                                                                If unea_panels_count = "" Then msgbox "2249 Delete after testing -- unea_panels_count is blank, it will probably error" 
                                                                                unea_panels_count = unea_panels_count * 1

                                                                                'If the loop has reached the final panel without encountering a Type 36 message, then it updates the processing notes accordingly
                                                                                If panel_count = unea_panels_count Then
                                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                    Exit Do
                                                                                End If
                                                                            Loop
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If

                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                            'There is at least one missing Type 36 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                            'All of the identified HH members have a corresponding Type 36 UNEA panel. The message can be deleted.
                                                            list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            dail_row = dail_row - 1
                                                        End If


                                                    Else
                                                        'There are PMIs in the DAIL message that are not on the case. Therefore, this message should be flagged for QI and added to the DAIL skip list when it is encountered again.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    End If

                                                    'Navigate back to the DAIL. This will reset to the top of the DAIL messages for the specific case number. Need to consider how to handle.
                                                    PF3

                                                ElseIf InStr(dail_msg, "DISB SPOUSAL SUP (TYPE 37)") Then
                                                    'Reset the caregiver reference number
                                                    caregiver_ref_nbr = ""

                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                    EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                    EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                    EMReadScreen check_full_dail_msg_line_4, 60, 12, 5

                                                    If trim(check_full_dail_msg_line_2) = "" Then 
                                                        check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                    End If

                                                    check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- check_full_dail_msg <> full_dail_msg. STOP HERE. Something went wrong."
                                                    End if

                                                    'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "REF NBR:", row, col
                                                    EMReadScreen caregiver_ref_nbr, 2, row, col + 9

                                                    'Transmit back to DAIL message
                                                    transmit

                                                    'Navigate to STAT/UNEA to check for corresponding Type 37 UNEA panel
                                                    Call write_value_and_transmit("S", dail_row, 3)
                                                    Call write_value_and_transmit("UNEA", 20, 71)

                                                    'Open the first panel of the caregiver reference number
                                                    EMWriteScreen caregiver_ref_nbr, 20, 76
                                                    Call write_value_and_transmit("01", 20, 79)

                                                    'Check if no UNEA panel exists
                                                    EmReadScreen unea_panel_check, 25, 24, 2

                                                    'Check if UNEA panels exist for the caregiver reference number
                                                    If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                        'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for caregiver M" & caregiver_ref_nbr & ".")
                                                    Else
                                                        'Read the UNEA type
                                                        EMReadScreen unea_type, 2, 5, 37
                                                        If unea_type = "37" Then
                                                            'If it is a type 37 panel then it does not need to read the other panels
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."
                                                        Else
                                                            'Check how many panels exist for the HH member
                                                            EMReadScreen unea_panels_count, 1, 2, 78
                                                            
                                                            If unea_panels_count = "1" Then
                                                                'If there is only one UNEA panel and it is not a Type 37 then it will update processing notes
                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."

                                                            ElseIf unea_panels_count <> "1" Then
                                                                'If there are more than just a single UNEA panel, loop through them all to check for Type 37
                                                                'Set incrementor for do loop
                                                                panel_count = 1

                                                                Do
                                                                    panel_count = panel_count + 1
                                                                    EMWriteScreen caregiver_ref_nbr, 20, 76
                                                                    Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                    'Read the UNEA type
                                                                    EMReadScreen unea_type, 2, 5, 37
                                                                    If unea_type = "37" Then
                                                                        'If it is a type 36 panel then it does not need to read the other panels
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                        Exit Do
                                                                    End if

                                                                    'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                    panel_count = panel_count * 1
                                                                    unea_panels_count = unea_panels_count * 1

                                                                    'If the loop has reached the final panel without encountering a Type 37 message, then it updates the processing notes accordingly
                                                                    If panel_count = unea_panels_count Then
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                        Exit Do
                                                                    End If
                                                                Loop
                                                            End If
                                                        End If
                                                    End If


                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                            'There is at least one missing Type 37 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                            'All of the identified HH members have a corresponding Type 37 UNEA panel. The message can be deleted.
                                                            list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                            dail_row = dail_row - 1
                                                        End If

                                                    'PF3 back to DAIL
                                                    PF3

                                                ElseIf InStr(dail_msg, "DISB CS ARREARS (TYPE 39) OF") Then
                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                    EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                    EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                    EMReadScreen check_full_dail_msg_line_4, 60, 12, 5

                                                    If trim(check_full_dail_msg_line_2) = "" Then 
                                                        check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                    End If

                                                    check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- check_full_dail_msg <> full_dail_msg. STOP HERE. Something went wrong."
                                                    End if

                                                    ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                    'Identify where 'PMI(S):' text is so that script can account for Type 39 and replaced Type 39 is
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "PMI(S):", row, col
                                                    EMReadScreen PMIs_line_one, 65 - (col + 8), row, col + 8
                                                    EMReadScreen PMIs_line_two, 60, 11, 5
                                                    EMReadScreen PMIs_line_three, 60, 12, 5
                                                    
                                                    'Combine the PMIs into one string
                                                    full_PMIs = replace(PMIs_line_one & PMIs_line_two & PMIs_line_three, " ", "")
                                                    'Splits the PMIs into an array
                                                    PMIs_array = Split(full_PMIs, ",")

                                                    'Reset the array 
                                                    ReDim PMI_and_ref_nbr_array(3, 0)

                                                    'Using list of PMIs in PMIs_array to update the PMI number in PMI_and_ref_nbr_array 
                                                    for each_PMI = 0 to UBound(PMIs_array, 1)
                                                        ReDim Preserve PMI_and_ref_nbr_array(hh_member_count_const, each_PMI)
                                                        PMI_and_ref_nbr_array(PMI_const, each_PMI) = PMIs_array(each_PMI)
                                                    Next 

                                                    'Transmit back to DAIL
                                                    transmit

                                                    ' Navigate to CASE/PERS to match PMIs and Ref Nbrs for checking UNEA panel
                                                    Call write_value_and_transmit("H", dail_row, 3)

                                                    EMWriteScreen "PERS", 20, 69
                                                    Transmit

                                                    ' Iterate through CASE/PERS to match up PMI with Ref Nbr

                                                    'the first member number starts at row 10
                                                    pers_row = 10                  

                                                    Do
                                                        'Reset reference number and PMI number so they don't carry through loop
                                                        ref_number_pers_panel = ""
                                                        pmi_number_pers_panel = ""

                                                        'reading the member number
                                                        EMReadScreen ref_number_pers_panel, 2, pers_row, 3
                                                        ref_number_pers_panel = trim(ref_number_pers_panel)

                                                        'Reading the PMI number
                                                        EMReadScreen pmi_number_pers_panel, 8, pers_row, 34  
                                                        pmi_number_pers_panel = trim(pmi_number_pers_panel)

                                                        for each_PMI = 0 to UBound(PMI_and_ref_nbr_array, 2)

                                                            If pmi_number_pers_panel = PMI_and_ref_nbr_array(PMI_const, each_PMI) Then
                                                                PMI_and_ref_nbr_array(ref_nbr_const, each_PMI) = ref_number_pers_panel
                                                                PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI) = True
                                                            End If
                                                        Next
                                                        
                                                        'go to the next member number - which is 3 rows down
                                                        pers_row = pers_row + 3

                                                        'if it reaches 19 - this is further down from the last member
                                                        If pers_row = 19 Then                       
                                                            'go to the next page and reset to line 10
                                                            PF8         
                                                            EMReadScreen last_page_check, 21, 24, 2                          
                                                            If last_page_check = "THIS IS THE LAST PAGE" Then Exit Do   
                                                            pers_row = 10
                                                        End If

                                                        EMReadScreen ref_number_pers_panel, 2, pers_row, 3
                                                    Loop until ref_number_pers_panel = "  "      
                                                    
                                                    'If there are PMIs listed on the DAIL message that do not align, then that DAIL message needs to be flagged for QI
                                                    for each_individual = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                        If PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) <> True Then
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " not found on case.")
                                                        ElseIf PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) = True Then
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " found on case (M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ").")
                                                        Else
                                                            MsgBox "Testing -- Something went wrong with matching PMIs for Type 39"
                                                        End If
                                                    Next

                                                    'Only check UNEA panels if ALL PMIs are matched for DAIL message and for case. There are no PMIs that did not match within the array.
                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case") = 0 Then
                                                        'If all PMIs are found on the case, then the script will navigate directly to STAT/UNEA from CASE/PERS to verify that UNEA panels exist for CS Type 39 for each identified PMI/reference number

                                                        'Update the processing notes to indicate that all PMIs found on the case rather than listing out on by one
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = "All PMIs found on case. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        EMWriteScreen "STAT", 19, 22
                                                        Call write_value_and_transmit("UNEA", 19, 69)

                                                        EmReadScreen no_unea_panels_exist, 34, 24, 2

                                                        If trim(no_unea_panels_exist) = "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                            'If no UNEA panels exist for the case, then the case needs to be flagged for QI
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = " No UNEA panels exist for any member on the case." & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        ElseIf trim(no_unea_panels_exist) <> "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                            'There are at least some UNEA panels. Iterate through all of the PMI/reference numbers to ensure there are corresponding UNEA panels for the DISB Type
                                                            for each_individual = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                                'Navigate to first UNEA panel for member to determine if any exist
                                                                EMWriteScreen PMI_and_ref_nbr_array(ref_nbr_const, each_individual), 20, 76
                                                                Call write_value_and_transmit("01", 20, 79)

                                                                'Check if no UNEA panel exists
                                                                EmReadScreen unea_panel_check, 25, 24, 2

                                                                If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                                    'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ".")
                                                                Else
                                                                    'Read the UNEA type
                                                                    EMReadScreen unea_type, 2, 5, 37
                                                                    If unea_type = "39" Then
                                                                        'If it is a type 39 panel then it does not need to read the other panels
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                    Else
                                                                        'Check how many panels exist for the HH member
                                                                        EMReadScreen unea_panels_count, 1, 2, 78

                                                                        If unea_panels_count = "1" Then
                                                                            'If there is only one UNEA panel and it is not a Type 39 then it will update processing notes
                                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                            
                                                                            
                                                                        ElseIf unea_panels_count <> "1" Then
                                                                            'If there are more than just a single UNEA panel, loop through them all to check for Type 39
                                                                            
                                                                            'Set incrementor for do loop
                                                                            panel_count = 1

                                                                            Do
                                                                                panel_count = panel_count + 1
                                                                                EMWriteScreen PMI_and_ref_nbr_array(ref_nbr_const, each_individual), 20, 76
                                                                                Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                                'Read the UNEA type
                                                                                EMReadScreen unea_type, 2, 5, 37
                                                                                If unea_type = "39" Then
                                                                                    'If it is a type 39 panel then it does not need to read the other panels
                                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                    Exit Do
                                                                                End if

                                                                                'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                                panel_count = panel_count * 1
                                                                                unea_panels_count = unea_panels_count * 1

                                                                                'If the loop has reached the final panel without encountering a Type 39 message, then it updates the processing notes accordingly
                                                                                If panel_count = unea_panels_count Then
                                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                    Exit Do
                                                                                End If
                                                                            Loop
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If

                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                            'There is at least one missing Type 39 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                            'All of the identified HH members have a corresponding Type 39 UNEA panel. The message can be deleted.
                                                            list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                            dail_row = dail_row - 1
                                                        End If


                                                    Else
                                                        'There are PMIs in the DAIL message that are not on the case. Therefore, this message should be flagged for QI and added to the DAIL skip list when it is encountered again.

                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1

                                                    End If

                                                    'Navigate back to the DAIL. This will reset to the top of the DAIL messages for the specific case number. Need to consider how to handle.
                                                    PF3

                                                ElseIf InStr(dail_msg, "DISB SPOUSAL SUP ARREARS (TYPE 40) OF") Then
                                                    'Reset the caregiver reference number
                                                    caregiver_ref_nbr = ""

                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                    EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                    EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                    EMReadScreen check_full_dail_msg_line_4, 60, 12, 5

                                                    If trim(check_full_dail_msg_line_2) = "" Then 
                                                        check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                    End If

                                                    check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- check_full_dail_msg <> full_dail_msg. STOP HERE. Something went wrong."
                                                    End if

                                                    'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "REF NBR:", row, col
                                                    EMReadScreen caregiver_ref_nbr, 2, row, col + 9

                                                    'Transmit back to DAIL message
                                                    transmit

                                                    'Navigate to STAT/UNEA to check for corresponding Type 37 UNEA panel
                                                    Call write_value_and_transmit("S", dail_row, 3)
                                                    Call write_value_and_transmit("UNEA", 20, 71)

                                                    'Open the first panel of the caregiver reference number
                                                    EMWriteScreen caregiver_ref_nbr, 20, 76
                                                    Call write_value_and_transmit("01", 20, 79)

                                                    'Check if no UNEA panel exists
                                                    EmReadScreen unea_panel_check, 25, 24, 2

                                                    'Check if UNEA panels exist for the caregiver reference number
                                                    If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                        'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for caregiver M" & caregiver_ref_nbr & ".")
                                                    Else
                                                        'Read the UNEA type
                                                        EMReadScreen unea_type, 2, 5, 37
                                                        If unea_type = "40" Then
                                                            'If it is a type 40 panel then it does not need to read the other panels
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                        Else
                                                            'Check how many panels exist for the HH member
                                                            EMReadScreen unea_panels_count, 1, 2, 78
                                                            If unea_panels_count = "1" Then
                                                                'If there is only one UNEA panel and it is not a Type 40 then it will update processing notes
                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "No UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                            
                                                            ElseIf unea_panels_count <> "1" Then
                                                                'If there are more than just a single UNEA panel, loop through them all to check for Type 40
                                                                'Set incrementor for do loop
                                                                panel_count = 1

                                                                Do
                                                                    panel_count = panel_count + 1
                                                                    EMWriteScreen caregiver_ref_nbr, 20, 76
                                                                    Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                    'Read the UNEA type
                                                                    EMReadScreen unea_type, 2, 5, 37
                                                                    If unea_type = "40" Then
                                                                        'If it is a type 40 panel then it does not need to read the other panels
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                        Exit Do
                                                                    End if

                                                                    'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                    panel_count = panel_count * 1
                                                                    unea_panels_count = unea_panels_count * 1
                                                                    
                                                                    'If the loop has reached the final panel without encountering a Type 40 message, then it updates the processing notes accordingly
                                                                    If panel_count = unea_panels_count Then
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "No UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                        Exit Do
                                                                    End If
                                                                Loop
                                                            End If
                                                        End If
                                                    End If

                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                            'There is at least one missing Type 40 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                            'All of the identified HH members have a corresponding Type 40 UNEA panel. The message can be deleted.
                                                            list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                            dail_row = dail_row - 1
                                                        End If

                                                    'PF3 back to DAIL
                                                    PF3

                                                ElseIf InStr(dail_msg, "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR:") Then
                                                    'Activate testing msgboxes here
                                                    ' activate_msg_boxes = True

                                                    If activate_msg_boxes = True then MsgBox "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: - In-scope message, evaluate how it works!"
                                                    activate_msg_boxes = False

                                                    'Reset variables
                                                    caregiver_ref_nbr = ""
                                                    no_exact_JOBS_panel_matches = ""
                                                    list_of_employers_on_jobs_panels = "*"
                                                    JOBS_footer_month = ""
                                                    JOBS_footer_year = ""
                                                    date_hired = ""
                                                    dail_msg = ""
                                                    check_full_dail_msg = ""
                                                    employer_full_name = ""
                                                    check_full_dail_msg = ""
                                                    case_name_to_return = ""

                                                    'To ensure script can get back to DAIL message if it gets bumped back to SELF, need to read the case name
                                                    EmReadScreen case_name_to_return, 3, dail_row - 1, 5
                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    'Check if the full message is displayed
                                                    EMReadScreen full_message_check, 36, 24, 2
                                                    If InStr(full_message_check, "THE ENTIRE MESSAGE TEXT") Then
                                                        EMReadScreen dail_msg, 61, dail_row, 20
                                                        dail_msg = trim(dail_msg)
                                                        check_full_dail_msg = dail_msg

                                                        'Since the entire message is displayed, script reads the reference number and employer name from the dail_msg string
                                                        caregiver_ref_nbr = Mid(check_full_dail_msg, instr(check_full_dail_msg, "REF NBR: ") + 9, 2)
                                                        employer_full_name = Mid(check_full_dail_msg, instr(check_full_dail_msg, "REF NBR: ") + 12, 8)
                                                        If activate_msg_boxes = True then MsgBox "caregiver_ref_nbr: " & caregiver_ref_nbr & "     employer_full_name: " & employer_full_name

                                                        'Remove x from dail message
                                                        EMWriteScreen " ", dail_row, 3
                                                    Else
                                                        ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                        
                                                        EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                        EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                        EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                        EMReadScreen check_full_dail_msg_line_4, 60, 12, 5

                                                        If trim(check_full_dail_msg_line_2) = "" Then check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)

                                                        check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                        'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                        'Set row and col
                                                        row = 1
                                                        col = 1
                                                        EMSearch "REF NBR:", row, col
                                                        EMReadScreen caregiver_ref_nbr, 2, row, col + 9

                                                        'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                        'Set row and col
                                                        row = 1
                                                        col = 1
                                                        EMSearch "REF NBR:", row, col
                                                        EMReadScreen employer_name_line_1, 8, row, col + 12

                                                        If trim(check_full_dail_msg_line_2) = "" Then 
                                                            employer_name_line_1 = trim(employer_name_line_1)
                                                        End If
                                                    
                                                        employer_full_name = trim(employer_name_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                        If activate_msg_boxes = True then MsgBox "caregiver_ref_nbr: " & caregiver_ref_nbr & "     employer_full_name: " & employer_full_name
                                                        
                                                        'Transmit back to DAIL message
                                                        transmit

                                                    End If

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- check_full_dail_msg <> full_dail_msg. STOP HERE. Something went wrong."
                                                    End if

                                                    'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
                                                    employer_full_name_split = split(employer_full_name, " ")

                                                    If len(employer_full_name_split(0)) < 4 and Ubound(employer_full_name_split) > 0 Then
                                                        employer_full_name_first_word = employer_full_name_split(0) & " " & employer_full_name_split(1)
                                                        If activate_msg_boxes = True then MsgBox "First word less than 3 characters long. employer_full_name_first_word is " & employer_full_name_first_word  
                                                    Else
                                                        employer_full_name_first_word = employer_full_name_split(0)   
                                                        If activate_msg_boxes = True then MsgBox "First word longer than 3 characters long. employer_full_name_first_word is " & employer_full_name_first_word
                                                    End If

                                                    If instr(len(employer_full_name_first_word), employer_full_name_first_word, ",") = len(employer_full_name_first_word) then 
                                                        employer_full_name_first_word = Mid(employer_full_name_first_word, 1, len(employer_full_name_first_word) - 1)
                                                        If activate_msg_boxes = True then MsgBox "Last character is a comma. HIRE_employer_name_first_word is now " & HIRE_employer_name_first_word
                                                    End If

                                                    'Navigate to STAT/JOBS to check if corresponding JOBS panel exists
                                                    Call write_value_and_transmit("S", dail_row, 3)
                                                    Call write_value_and_transmit("JOBS", 20, 71)

                                                    'Open the first JOBS panel of the caregiver reference number
                                                    EMWriteScreen caregiver_ref_nbr, 20, 76
                                                    Call write_value_and_transmit("01", 20, 79)
                                                    
                                                    'started adding in new JOBS code from function here
                                                    'Need to navigate to JOBS panel for CM if not there already so will check if at CM right now
                                                    EMReadScreen JOBS_footer_month_and_year, 5, 20, 55

                                                    If JOBS_footer_month_and_year <> CM_mo & " " & CM_yr then 
                                                        If activate_msg_boxes = True Then MsgBox "Testing -- Need to navigate to CM"
                                                        'PF3 back to DAIL and navigate to CASE/CURR to change the footer month and get to JOBS panel for CM
                                                        PF3
                                                        Call write_value_and_transmit("H", dail_row, 3)
                                                        EMReadScreen curr_panel_check, 4, 2, 55
                                                        If curr_panel_check <> "CURR" Then MsgBox "Testing -- not at CASE/CURR"
                                                        EMWriteScreen "STAT", 20, 22
                                                        EMWriteScreen CM_mo, 20, 54
                                                        EMWriteScreen CM_yr, 20, 57
                                                        Call write_value_and_transmit("JOBS", 20, 69)

                                                        'Open the first JOBS panel of the caregiver reference number
                                                        EMWriteScreen caregiver_ref_nbr, 20, 76
                                                        Call write_value_and_transmit("01", 20, 79)
                                                    Else    
                                                        If activate_msg_boxes = True Then MsgBox "Testing -- Already at CM JOBS panel"
                                                    End If

                                                    'Ensure we are on JOBS panel
                                                    EmReadScreen jobs_panel_nav_check, 4, 2, 45
                                                    If jobs_panel_nav_check <> "JOBS" Then MsgBox "Testing -- Not on JOBS panel. Stop here"
                                                    
                                                    If activate_msg_boxes = True Then MsgBox "Testing -- Ensure that we are on correct HH Member. Should be at HH Member: " & caregiver_ref_nbr

                                                    'Check if no JOBS panel exists on HH Memb JOBS panel
                                                    EmReadScreen jobs_panel_check, 25, 24, 2

                                                    If InStr(jobs_panel_check, "DOES NOT EXIST") Then
                                                        'There are no JOBS panels for this HH member. The script will add a new JOBS panel for the member
                                                        If activate_msg_boxes = True Then MsgBox "Testing -- No JOBS panel exist. Script will create new panel and fill it out. STOP HERE if needed in production."

                                                        Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

                                                        'Validation to ensure that script is able to open a new JOBS panel
                                                        EmReadScreen panel_count_plus_one_check, 1, 2, 73
                                                        panel_count_plus_one_check = panel_count_plus_one_check * 1
                                                        EmReadScreen panel_count_total_check, 1, 2, 78
                                                        panel_count_total_check = panel_count_total_check * 1

                                                        If panel_count_plus_one_check <> panel_count_total_check + 1 then 
                                                            If activate_msg_boxes = True Then MsgBox "Testing -- unable to open a new JOBS panel. Will note in spreadsheet and continue"
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "MAXIS programs are inactive. Unable to add a new JOBS panel for M" & caregiver_ref_nbr & ". Review needed." & " Message should not be deleted."
                                                        Else
                                                            
                                                            If activate_msg_boxes = True Then MsgBox "Testing -- Script opened JOBS panel. Will add new panel"

                                                            'Reads footer month for updating the panel
                                                            EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                            EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                            'No date hired from HIRE message so use dail_month
                                                            EmWriteScreen left(dail_month, 2), 9, 35
                                                            EMWriteScreen "01", 9, 38
                                                            EmWriteScreen right(dail_month, 2), 9, 41
                                                            
                                                            'Writes information to JOBS panel
                                                            EMWriteScreen "O", 5, 34
                                                            EMWriteScreen "4", 6, 34
                                                            EMWriteScreen employer_full_name, 7, 42
                                                                
                                                            'Script writes panel footer month and date to the new panel
                                                            EmWriteScreen JOBS_footer_month, 12, 54
                                                            EMWriteScreen "01", 12, 57
                                                            EmWriteScreen JOBS_footer_year, 12, 60

                                                            'Puts $0 in as the received income amt and 0 worked hours
                                                            EMWriteScreen "0", 12, 67				
                                                            EMWriteScreen "0", 18, 72	
                                                            
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review the JOBS panel. Any potential errors or issues before it continues?"
                                                            
                                                            'Opens FS PIC
                                                            Call write_value_and_transmit("X", 19, 38)
                                                                
                                                            'Write today's date to calculation since added today
                                                            Call create_MAXIS_friendly_date(date, 0, 5, 34)
                                                            
                                                            'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                                                            EMWriteScreen "1", 5, 64
                                                            EMWriteScreen "0", 8, 64
                                                            EMWriteScreen "0", 9, 66
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review the PIC panel. Any potential errors or issues before it continues?"

                                                            transmit
                                                            EmReadScreen PIC_warning, 7, 20, 6
                                                            IF PIC_warning = "WARNING" then transmit 'to clear message
                                                            transmit 'back to JOBS panel
                                                            If activate_msg_boxes = True Then Msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                                                            transmit 'to save JOBS panel

                                                            'Check if information is expiring and needs to be added to a future month
                                                            EMReadScreen expired_check, 6, 24, 17 
                                                            EMReadScreen data_expiration_month, 2, 24, 27
                                                            EMReadScreen jobs_panel_month, 2, 20, 55

                                                            If expired_check = "EXPIRE" THEN 
                                                                Do
                                                                    'Do loop to add JOBS panels to every month from DAIL month through CM
                                                                    If activate_msg_boxes = True Then msgbox "Testing -- New JOBS panel is expiring so it needs to be added to CM + 1 as well"

                                                                    'PF3 to go to STAT/WRAP
                                                                    PF3

                                                                    'Check to make sure on STAT/WRAP
                                                                    EMReadScreen stat_wrap_check, 19, 2, 32
                                                                    If Instr(stat_wrap_check, "Wrap") = 0 Then MsgBox "Testing -- It didn't go to STAT/WRAP for some reason. Stop here!!"

                                                                    'Build do loop to get to expiration month so that it isn't creating a bunch of duplicate JOBS panels
                                                                    If data_expiration_month <> jobs_panel_month Then
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- JOBS panel expires in future month"
                                                                        Do
                                                                            Call write_value_and_transmit("Y", 16, 54)
                                                                            EMReadScreen stat_wrap_month_check, 2, 20, 55
                                                                            If stat_wrap_month_check = data_expiration_month Then
                                                                                'Script has reached the expiration month, it will go to next month and then exit
                                                                                If activate_msg_boxes = True Then msgbox "Testing -- script has found matching month"
                                                                                PF3
                                                                                Call write_value_and_transmit("Y", 16, 54)
                                                                                Exit Do
                                                                            Else
                                                                                'Script has not yet reached the expiration month, it will PF3 back to STAT/WRAP to move to next month
                                                                                If activate_msg_boxes = True Then msgbox "Testing -- script has not reached matching month"
                                                                                PF3
                                                                            End If
                                                                        Loop
                                                                    Else
                                                                        'If the expiration month and the jobs panel month are the same then it should add to next month too since it will expire at end of month
                                                                        Call write_value_and_transmit("Y", 16, 54)
                                                                    End If

                                                                    'Navigate to STAT/JOBS
                                                                    Call write_value_and_transmit("JOBS", 20, 71)

                                                                    EMReadScreen jobs_panel_nav_check, 8, 2, 43
                                                                    If InStr(jobs_panel_nav_check, "JOBS") = 0 Then MsgBox "Testing -- Stop here. Not at JOBS panel"

                                                                    If activate_msg_boxes = True Then MsgBox "Testing -- Is it at the month after expiration? Expiration month was " & data_expiration_month

                                                                    'Navigate to HH member
                                                                    Call write_value_and_transmit(caregiver_ref_nbr, 20, 76)

                                                                    'Making sure there aren't 5 jobs already
                                                                    EMReadScreen five_jobs_check, 1, 2, 78
                                                                    
                                                                    If five_jobs_check = "5" Then 
                                                                        script_end_procedure_with_error_report("Testing -- There are 5 JOBS panels already, it will error out. Must stop here!")
                                                                    Else
                                                                        Call write_value_and_transmit("NN", 20, 79)
                                                                    End If
                                                                    
                                                                    EmReadScreen panel_count_plus_one_check, 1, 2, 73
                                                                    panel_count_plus_one_check = panel_count_plus_one_check * 1
                                                                    EmReadScreen panel_count_total_check, 1, 2, 78
                                                                    panel_count_total_check = panel_count_total_check * 1

                                                                    If panel_count_plus_one_check <> panel_count_total_check + 1 then script_end_procedure_with_error_report("Testing -- Unable to open a new JOBS panel. Script will stop here.")

                                                                    'Reads footer month for updating the panel
                                                                    EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                                    EMReadScreen JOBS_footer_year, 2, 20, 58

                                                                    'No date hired from HIRE message so use dail_month
                                                                    EmWriteScreen left(dail_month, 2), 9, 35
                                                                    EMWriteScreen "01", 9, 38
                                                                    EmWriteScreen right(dail_month, 2), 9, 41

                                                                    'Writes information to JOBS panel
                                                                    EMWriteScreen "O", 5, 34
                                                                    EMWriteScreen "4", 6, 34
                                                                    EMWriteScreen employer_full_name, 7, 42
                                                                    
                                                                    'Looking at CM + 1 so won't match the message, just writes footer month to panel
                                                                    EmWriteScreen JOBS_footer_month, 12, 54
                                                                    EMWriteScreen "01", 12, 57
                                                                    EmWriteScreen JOBS_footer_year, 12, 60

                                                                    'Puts $0 in as the received income amt
                                                                    EMWriteScreen "0", 12, 67				
                                                                    'Puts 0 hours in as the worked hours
                                                                    EMWriteScreen "0", 18, 72		

                                                                    If activate_msg_boxes = True Then msgbox "Testing - Does everything look good on JOBS panel before heading to PIC?"
                                                                    
                                                                    'Opens FS PIC
                                                                    Call write_value_and_transmit("X", 19, 38)
                                                                    'Writes today's date on the panel
                                                                    Call create_MAXIS_friendly_date(date, 0, 5, 34)

                                                                    'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                                                                    EMWriteScreen "1", 5, 64
                                                                    EMWriteScreen "0", 8, 64
                                                                    EMWriteScreen "0", 9, 66
                                                                    If activate_msg_boxes = True Then msgbox "Testing - Does everything look good on JOBS panel before saving the PIC?"
                                                                    
                                                                    transmit
                                                                    EmReadScreen PIC_warning, 7, 20, 6
                                                                    IF PIC_warning = "WARNING" then transmit 'to clear message
                                                                    transmit 'back to JOBS panel
                                                                    If activate_msg_boxes = True Then msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                                                                    transmit 'to save JOBS panel
                                                                    
                                                                    'Check if information is expiring and needs to be added to CM + 1
                                                                    EMReadScreen expired_check, 6, 24, 17 
                                                                    EMReadScreen data_expiration_month, 2, 24, 27
                                                                    EMReadScreen jobs_panel_month, 2, 20, 55 
                                                                    
                                                                    If expired_check <> "EXPIRE" THEN
                                                                        'If data is not expiring, then the script can exit the do loop
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- No expiration date. It will exit the do loop"
                                                                        Exit Do
                                                                    Else
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- Data is expiring. It will continue with the do loop"
                                                                    End If

                                                                Loop

                                                            End If

                                                            'Write information to CASE/NOTE
                                                            If activate_msg_boxes = True Then MsgBox "Testing -- Script will now CASE/NOTE information. Navigate to CASE/NOTE"

                                                            'PF4 to navigate to CASE/NOTE
                                                            PF4

                                                            EMReadScreen jobs_panel_not_saved, 25, 24, 2
                                                            'If unable to navigate to CASE/NOTE due to not saving JOBS panel, then another transmit is needed
                                                            If instr(jobs_panel_not_saved, "CASE OR PERSON NOTES ARE") Then 
                                                                transmit
                                                                PF4
                                                            End If

                                                            EMReadScreen case_note_check, 4, 2, 45
                                                            If case_note_check <> "NOTE" then MsgBox "Testing -- not at case note stop here3"

                                                            'Open new CASE/NOTE
                                                            PF9

                                                            'Write information regarding CS EMPLOYER REPORTED match
                                                            CALL write_variable_in_case_note("-CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: " & caregiver_ref_nbr & " for " & trim(employer_full_name) & "-")
                                                            CALL write_variable_in_case_note("DAIL MONTH: " & dail_month)
                                                            CALL write_variable_in_case_note("EMPLOYER: " & employer_full_name)
                                                            CALL write_variable_in_case_note("---")
                                                            CALL write_variable_in_case_note("NO CORRESPONDING JOBS PANEL EXISTED FOR EMPLOYER NOTED IN CSES MESSAGE. STAT/JOBS PANEL ADDED FOR EMPLOYER IDENTIFIED IN CSES DAIL MESSAGE.")
                                                            CALL write_variable_in_case_note("---")
                                                            CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                                                            CALL write_variable_in_case_note("---")
                                                            CALL write_variable_in_case_note(worker_signature)

                                                            If activate_msg_boxes = True Then msgbox "Testing -- The script is about to save the CASE/NOTE. Stop here if in testing or production"

                                                            'PF3 to save the CASE/NOTE
                                                            PF3
                                                            
                                                            'PF3 to STAT/WRAP or JOBS
                                                            PF3
                                                            
                                                            EMReadScreen panel_nav_check, 4, 2, 46
                                                            If panel_nav_check <> "WRAP" Then
                                                                PF3
                                                                If activate_msg_boxes = True Then msgbox "Testing -- The script should now be at STAT/WRAP. If it is not, then stop here."
                                                            End If

                                                            If activate_msg_boxes = True Then msgbox "Testing -- No jobs panels existed. Created JOBS panel(s) through CM"
                                                            
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No JOBS panels exist for household member number: " & caregiver_ref_nbr & ". JOBS Panel and CASE/NOTE added for employer noted in HIRE message. Message should be deleted.")
                                                        End If

                                                    Else
                                                        'There is at least 1 JOBS panel
                                                        If activate_msg_boxes = True Then MsgBox "Testing -- there is at least 1 JOBS panel."

                                                        'Read the employer name, but only first 20 characters to align with max length for HIRE message for NDNH messages
                                                        EMReadScreen employer_name_jobs_panel, 20, 7, 42
                                                        employer_name_jobs_panel = trim(replace(employer_name_jobs_panel, "_", " "))

                                                        'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
                                                        employer_name_jobs_panel_split = split(employer_name_jobs_panel, " ")

                                                        If len(employer_name_jobs_panel_split(0)) < 4 and Ubound(employer_name_jobs_panel_split) > 0 Then
                                                            employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0) & " " & employer_name_jobs_panel_split(1)
                                                            If activate_msg_boxes = True Then MsgBox "First word less than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word  
                                                        Else
                                                            employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0)   
                                                            If activate_msg_boxes = True Then MsgBox "First word longer than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word
                                                        End If

                                                        If instr(len(employer_name_jobs_panel_first_word), employer_name_jobs_panel_first_word, ",") = len(employer_name_jobs_panel_first_word) then 
                                                            employer_name_jobs_panel_first_word = Mid(employer_name_jobs_panel_first_word, 1, len(employer_name_jobs_panel_first_word) - 1)
                                                            If activate_msg_boxes = True Then MsgBox "Last character is a comma. employer_name_jobs_panel_first_word is now " & employer_name_jobs_panel_first_word
                                                        End If

                                                        If employer_name_jobs_panel = employer_full_name Then
                                                            'Add here
                                                            If activate_msg_boxes = True Then msgbox "Testing -- The employer names match exactly. Will add to delete list and TIKL delete list."

                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". JOBS panel matches HIRE employer name exactly. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 

                                                        ElseIf employer_name_jobs_panel_first_word = employer_full_name_first_word Then

                                                            If activate_msg_boxes = True Then msgbox "Testing -- there is an exact match for employer name first word only. Will add to delete list and TIKL delete list."

                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". JOBS panel matches first word of HIRE employer name. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 

                                                        Else
                                                            'Check how many panels exist for the HH member
                                                            EMReadScreen jobs_panels_count, 1, 2, 78
                                                            'Convert jobs_panels_count to a number
                                                            jobs_panels_count = jobs_panels_count * 1
                                                            'If there is more than just 1 JOBS panel, loop through them all to check for matching employers
                                                            If jobs_panels_count = 1 Then
                                                                If activate_msg_boxes = True Then MsgBox "Testing -- There is only one JOBS panel and they do not match. The script will skip the message since there is no exact match"

                                                                'Set variable below to true to trigger dialog
                                                                no_exact_JOBS_panel_matches = True
                                                            
                                                            ElseIf jobs_panels_count <> 1 Then
                                                                If activate_msg_boxes = True Then MsgBox "Testing -- There are multiple JOBS panels. Script will determine if there are any perfect matches."
                                                                
                                                                'Set incrementor for do loop
                                                                panel_count = 1

                                                                Do
                                                                    panel_count = panel_count + 1
                                                                    EMWriteScreen HIRE_memb_number, 20, 76
                                                                    Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                    'Read the employer name
                                                                    EMReadScreen employer_name_jobs_panel, 20, 7, 42
                                                                    employer_name_jobs_panel = trim(replace(employer_name_jobs_panel, "_", " "))

                                                                    'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
                                                                    employer_name_jobs_panel_split = split(employer_name_jobs_panel, " ")

                                                                    If len(employer_name_jobs_panel_split(0)) < 4 and Ubound(employer_name_jobs_panel_split) > 0 Then
                                                                        employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0) & " " & employer_name_jobs_panel_split(1)
                                                                        If activate_msg_boxes = True Then MsgBox "First word less than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word  
                                                                    Else
                                                                        employer_name_jobs_panel_first_word = employer_name_jobs_panel_split(0)   
                                                                        If activate_msg_boxes = True Then MsgBox "First word longer than 3 characters long. employer_name_jobs_panel_first_word is " & employer_name_jobs_panel_first_word
                                                                    End If

                                                                    If instr(len(employer_name_jobs_panel_first_word), employer_name_jobs_panel_first_word, ",") = len(employer_name_jobs_panel_first_word) then 
                                                                        employer_name_jobs_panel_first_word = Mid(employer_name_jobs_panel_first_word, 1, len(employer_name_jobs_panel_first_word) - 1)
                                                                        If activate_msg_boxes = True Then MsgBox "Last character is a comma. employer_name_jobs_panel_first_word is now " & employer_name_jobs_panel_first_word
                                                                    End If

                                                                    If employer_name_jobs_panel = employer_full_name Then
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- The employer names match exactly. Will add to delete list and TIKL delete list."
                                                            
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". JOBS panel matches HIRE employer name exactly. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 
                                                            
                                                                        'Exit the do loop since an exact match was found
                                                                        Exit Do
                                                            
                                                                    ElseIf employer_name_jobs_panel_first_word = employer_full_name_first_word Then
                                                            
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- there is an exact match for employer name first word only. Will add to delete list and TIKL delete list."
                                                            
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". JOBS panel matches first word of HIRE employer name. No CASE/NOTE created. Created TIKLs should be removed. Message should be deleted." 
                                                            
                                                                        'Exit the do loop since an exact match was found
                                                                        Exit Do

                                                                    End If

                                                                    'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                    panel_count = panel_count * 1
                                                                    jobs_panels_count = jobs_panels_count * 1
                                                                    
                                                                    If panel_count = jobs_panels_count Then
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- Since there were no exact employer matches, setting no_exact_JOBS_panel_matches = True"
                                                                        'Since there were no exact employer matches, setting no_exact_JOBS_panel_matches = True
                                                                        no_exact_JOBS_panel_matches = True
                                                                        Exit Do
                                                                    End If
                                                                Loop
                                                            End If

                                                            'Convert string of the employer names into an array for use in the dialog
                                                            If no_exact_JOBS_panel_matches = True Then

                                                                'If there are 5 jobs already, it will not add another JOBS panel
                                                                If jobs_panels_count = 5 Then
                                                                    'Script will be unable to add another JOBS panel since there are 5 already so it will note as such and skip
                                                                    If activate_msg_boxes = True Then msgbox "Testing -- There are 5 JOBS panels already. Cannot add another JOBS panel."

                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "There does not appear to be an exactly matching JOBS panel for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". However, unable to add new JOBS panel for employer since there are already 5 JOBS panels. Review needed." & " Message should not be deleted."

                                                                Else
                                                                    'There are not 5 JOBS panels so it will check CM + 1 too before adding new JOBS panel before adding new JOBS panel
                                                                    If activate_msg_boxes = True Then MsgBox "Testing -- Need to navigate to CM + 1"
                                                                    'PF3 back to DAIL and navigate to CASE/CURR to change the footer month and get to JOBS panel for CM
                                                                    PF3
                                                                    Call write_value_and_transmit("H", dail_row, 3)
                                                                    EMReadScreen curr_panel_check, 4, 2, 55
                                                                    If curr_panel_check <> "CURR" Then MsgBox "Testing -- not at CASE/CURR"
                                                                    EMWriteScreen "STAT", 20, 22
                                                                    EMWriteScreen CM_plus_1_mo, 20, 54
                                                                    EMWriteScreen CM_plus_1_yr, 20, 57
                                                                    Call write_value_and_transmit("JOBS", 20, 69)

                                                                    'Open the first JOBS panel of the caregiver reference number
                                                                    EMWriteScreen caregiver_ref_nbr, 20, 76
                                                                    Call write_value_and_transmit("01", 20, 79)

                                                                    'Read the number of JOBS panels to ensure there are not 5 already
                                                                    EMReadScreen jobs_panels_count_CM_plus_1, 1, 2, 78

                                                                    If jobs_panels_count_CM_plus_1 = "5" Then
                                                                        If activate_msg_boxes = True Then msgbox "Testing -- There are 5 JOBS panels already in CM + 1. Cannot add another JOBS panel."

                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "There does not appear to be an exactly matching JOBS panel for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". However, unable to add new JOBS panel for employer since there are already 5 JOBS panels in CM + 1. Review needed." & " Message should not be deleted."
                                                                    Else
                                                                        'Navigate back to CM and add JOBS as originally intended
                                                                        'There are not 5 JOBS panels so it will check CM + 1 too before adding new JOBS panel before adding new JOBS panel
                                                                        If activate_msg_boxes = True Then MsgBox "Testing -- Need to navigate to CM + 1"
                                                                        'PF3 back to DAIL and navigate to CASE/CURR to change the footer month and get to JOBS panel for CM
                                                                        PF3
                                                                        Call write_value_and_transmit("H", dail_row, 3)
                                                                        EMReadScreen curr_panel_check, 4, 2, 55
                                                                        If curr_panel_check <> "CURR" Then MsgBox "Testing -- not at CASE/CURR"
                                                                        EMWriteScreen "STAT", 20, 22
                                                                        EMWriteScreen CM_mo, 20, 54
                                                                        EMWriteScreen CM_yr, 20, 57
                                                                        Call write_value_and_transmit("JOBS", 20, 69)

                                                                        'Open the first JOBS panel of the caregiver reference number
                                                                        EMWriteScreen caregiver_ref_nbr, 20, 76
                                                                        Call write_value_and_transmit("01", 20, 79)

                                                                        Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

                                                                        'Validation to ensure that script is able to open a new JOBS panel
                                                                        EmReadScreen panel_count_plus_one_check, 1, 2, 73
                                                                        panel_count_plus_one_check = panel_count_plus_one_check * 1
                                                                        EmReadScreen panel_count_total_check, 1, 2, 78
                                                                        panel_count_total_check = panel_count_total_check * 1

                                                                        If panel_count_plus_one_check <> panel_count_total_check + 1 then 
                                                                            If activate_msg_boxes = True Then MsgBox "Testing -- unable to open a new JOBS panel. Will note in spreadsheet and continue"
                                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "MAXIS programs are inactive. Unable to add a new JOBS panel for M" & caregiver_ref_nbr & ". Review needed." & " Message should not be deleted."
                                                                        Else
                                                                            
                                                                            If activate_msg_boxes = True Then MsgBox "Testing -- Script opened JOBS panel. Will add new panel"

                                                                            'Reads footer month for updating the panel
                                                                            EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                                            EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                                            'No date hired from HIRE message so use dail_month
                                                                            EmWriteScreen left(dail_month, 2), 9, 35
                                                                            EMWriteScreen "01", 9, 38
                                                                            EmWriteScreen right(dail_month, 2), 9, 41

                                                                            'Writes information to JOBS panel
                                                                            EMWriteScreen "O", 5, 34
                                                                            EMWriteScreen "4", 6, 34
                                                                            EMWriteScreen employer_full_name, 7, 42
                                                                                
                                                                            'Otherwise, write the panel footer month and date to the new panel
                                                                            EmWriteScreen JOBS_footer_month, 12, 54
                                                                            EMWriteScreen "01", 12, 57
                                                                            EmWriteScreen JOBS_footer_year, 12, 60

                                                                            'Puts $0 in as the received income amt and 0 worked hours
                                                                            EMWriteScreen "0", 12, 67				
                                                                            EMWriteScreen "0", 18, 72	
                                                                            
                                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review the JOBS panel. Any potential errors or issues before it continues?"
                                                                            
                                                                            'Opens FS PIC
                                                                            Call write_value_and_transmit("X", 19, 38)
                                                                                
                                                                            'Write today's date to calculation since added today
                                                                            Call create_MAXIS_friendly_date(date, 0, 5, 34)
                                                                            
                                                                            'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                                                                            EMWriteScreen "1", 5, 64
                                                                            EMWriteScreen "0", 8, 64
                                                                            EMWriteScreen "0", 9, 66
                                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review the PIC panel. Any potential errors or issues before it continues?"

                                                                            transmit
                                                                            EmReadScreen PIC_warning, 7, 20, 6
                                                                            IF PIC_warning = "WARNING" then transmit 'to clear message
                                                                            transmit 'back to JOBS panel
                                                                            If activate_msg_boxes = True Then Msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                                                                            transmit 'to save JOBS panel

                                                                            'Check if information is expiring and needs to be added to a future month
                                                                            EMReadScreen expired_check, 6, 24, 17 
                                                                            EMReadScreen data_expiration_month, 2, 24, 27
                                                                            EMReadScreen jobs_panel_month, 2, 20, 55

                                                                            If expired_check = "EXPIRE" THEN 
                                                                                Do
                                                                                    'Do loop to add JOBS panels to every month from DAIL month through CM
                                                                                    If activate_msg_boxes = True Then msgbox "Testing -- New JOBS panel is expiring so it needs to be added to CM + 1 as well"

                                                                                    'PF3 to go to STAT/WRAP
                                                                                    PF3

                                                                                    'Check to make sure on STAT/WRAP
                                                                                    EMReadScreen stat_wrap_check, 19, 2, 32
                                                                                    If Instr(stat_wrap_check, "Wrap") = 0 Then MsgBox "Testing -- It didn't go to STAT/WRAP for some reason. Stop here!!"

                                                                                    'Build do loop to get to expiration month so that it isn't creating a bunch of duplicate JOBS panels
                                                                                    If data_expiration_month <> jobs_panel_month Then
                                                                                        If activate_msg_boxes = True Then msgbox "Testing -- JOBS panel expires in future month"
                                                                                        Do
                                                                                            Call write_value_and_transmit("Y", 16, 54)
                                                                                            EMReadScreen stat_wrap_month_check, 2, 20, 55
                                                                                            If stat_wrap_month_check = data_expiration_month Then
                                                                                                'Script has reached the expiration month, it will go to next month and then exit
                                                                                                If activate_msg_boxes = True Then msgbox "Testing -- script has found matching month"
                                                                                                PF3
                                                                                                Call write_value_and_transmit("Y", 16, 54)
                                                                                                Exit Do
                                                                                            Else
                                                                                                'Script has not yet reached the expiration month, it will PF3 back to STAT/WRAP to move to next month
                                                                                                If activate_msg_boxes = True Then msgbox "Testing -- script has not reached matching month"
                                                                                                PF3
                                                                                            End If
                                                                                        Loop
                                                                                    Else
                                                                                        'If the expiration month and the jobs panel month are the same then it should add to next month too since it will expire at end of month
                                                                                        Call write_value_and_transmit("Y", 16, 54)
                                                                                    End If

                                                                                    'Navigate to STAT/JOBS
                                                                                    Call write_value_and_transmit("JOBS", 20, 71)

                                                                                    EMReadScreen jobs_panel_nav_check, 8, 2, 43
                                                                                    If InStr(jobs_panel_nav_check, "JOBS") = 0 Then MsgBox "Testing -- Stop here. Not at JOBS panel"

                                                                                    If activate_msg_boxes = True Then MsgBox "Testing -- Is it at the month after expiration? Expiration month was " & data_expiration_month

                                                                                    'Navigate to HH member
                                                                                    Call write_value_and_transmit(caregiver_ref_nbr, 20, 76)

                                                                                    'Making sure there aren't 5 jobs already
                                                                                    EMReadScreen five_jobs_check, 1, 2, 78
                                                                                    
                                                                                    If five_jobs_check = "5" Then 
                                                                                        script_end_procedure_with_error_report("Testing -- There are 5 JOBS panels already, it will error out. Must stop here!")
                                                                                    Else
                                                                                        Call write_value_and_transmit("NN", 20, 79)
                                                                                    End If
                                                                                    
                                                                                    EmReadScreen panel_count_plus_one_check, 1, 2, 73
                                                                                    panel_count_plus_one_check = panel_count_plus_one_check * 1
                                                                                    EmReadScreen panel_count_total_check, 1, 2, 78
                                                                                    panel_count_total_check = panel_count_total_check * 1

                                                                                    If panel_count_plus_one_check <> panel_count_total_check + 1 then script_end_procedure_with_error_report("Testing -- Unable to open a new JOBS panel. Script will stop here.")

                                                                                    'Reads footer month for updating the panel
                                                                                    EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                                                    EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                                                    'No date hired from HIRE message so use dail_month
                                                                                    EmWriteScreen left(dail_month, 2), 9, 35
                                                                                    EMWriteScreen "01", 9, 38
                                                                                    EmWriteScreen right(dail_month, 2), 9, 41

                                                                                    'Writes information to JOBS panel
                                                                                    EMWriteScreen "O", 5, 34
                                                                                    EMWriteScreen "4", 6, 34
                                                                                    EMWriteScreen employer_full_name, 7, 42
                                                                                    
                                                                                    'Looking at CM + 1 so won't match the message, just writes footer month to panel
                                                                                    EmWriteScreen JOBS_footer_month, 12, 54
                                                                                    EMWriteScreen "01", 12, 57
                                                                                    EmWriteScreen JOBS_footer_year, 12, 60

                                                                                    'Puts $0 in as the received income amt
                                                                                    EMWriteScreen "0", 12, 67				
                                                                                    'Puts 0 hours in as the worked hours
                                                                                    EMWriteScreen "0", 18, 72		

                                                                                    If activate_msg_boxes = True Then msgbox "Testing - Does everything look good on JOBS panel before heading to PIC?"
                                                                                    
                                                                                    'Opens FS PIC
                                                                                    Call write_value_and_transmit("X", 19, 38)
                                                                                    'Writes today's date on the panel
                                                                                    Call create_MAXIS_friendly_date(date, 0, 5, 34)

                                                                                    'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                                                                                    EMWriteScreen "1", 5, 64
                                                                                    EMWriteScreen "0", 8, 64
                                                                                    EMWriteScreen "0", 9, 66
                                                                                    If activate_msg_boxes = True Then msgbox "Testing - Does everything look good on JOBS panel before saving the PIC?"
                                                                                    
                                                                                    transmit
                                                                                    EmReadScreen PIC_warning, 7, 20, 6
                                                                                    IF PIC_warning = "WARNING" then transmit 'to clear message
                                                                                    transmit 'back to JOBS panel
                                                                                    If activate_msg_boxes = True Then msgbox "Testing -- It is about save the JOBS panel. Stop here if in testing or production"
                                                                                    transmit 'to save JOBS panel
                                                                                    
                                                                                    'Check if information is expiring and needs to be added to CM + 1
                                                                                    EMReadScreen expired_check, 6, 24, 17 
                                                                                    EMReadScreen data_expiration_month, 2, 24, 27
                                                                                    EMReadScreen jobs_panel_month, 2, 20, 55 
                                                                                    
                                                                                    If expired_check <> "EXPIRE" THEN
                                                                                        'If data is not expiring, then the script can exit the do loop
                                                                                        If activate_msg_boxes = True Then msgbox "Testing -- No expiration date. It will exit the do loop"
                                                                                        Exit Do
                                                                                    Else
                                                                                        If activate_msg_boxes = True Then msgbox "Testing -- Data is expiring. It will continue with the do loop"
                                                                                    End If

                                                                                Loop

                                                                            End If

                                                                            'Write information to CASE/NOTE
                                                                            If activate_msg_boxes = True Then MsgBox "Testing -- Script will now CASE/NOTE information. Navigate to CASE/NOTE"

                                                                            'PF4 to navigate to CASE/NOTE
                                                                            PF4

                                                                            EMReadScreen jobs_panel_not_saved, 25, 24, 2
                                                                            'If unable to navigate to CASE/NOTE due to not saving JOBS panel, then another transmit is needed
                                                                            If instr(jobs_panel_not_saved, "CASE OR PERSON NOTES ARE") Then 
                                                                                transmit
                                                                                PF4
                                                                            End If

                                                                            EMReadScreen case_note_check, 4, 2, 45
                                                                            If case_note_check <> "NOTE" then MsgBox "Testing -- not at case note stop here4"

                                                                            'Open new CASE/NOTE
                                                                            PF9

                                                                            'Write information regarding CS EMPLOYER REPORTED match
                                                                            CALL write_variable_in_case_note("-CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: " & caregiver_ref_nbr & " for " & trim(employer_full_name) & "-")
                                                                            CALL write_variable_in_case_note("DAIL MONTH: " & dail_month)
                                                                            CALL write_variable_in_case_note("EMPLOYER: " & employer_full_name)
                                                                            CALL write_variable_in_case_note("---")
                                                                            CALL write_variable_in_case_note("NO CORRESPONDING JOBS PANEL EXISTED FOR EMPLOYER NOTED IN CSES MESSAGE. STAT/JOBS PANEL ADDED FOR EMPLOYER IDENTIFIED IN CSES DAIL MESSAGE.")
                                                                            CALL write_variable_in_case_note("---")
                                                                            CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                                                                            CALL write_variable_in_case_note("---")
                                                                            CALL write_variable_in_case_note(worker_signature)

                                                                            If activate_msg_boxes = True Then msgbox "Testing -- The script is about to save the CASE/NOTE. Stop here if in testing or production"

                                                                            'PF3 to save the CASE/NOTE
                                                                            PF3
                                                                            
                                                                            'PF3 to STAT/WRAP or JOBS
                                                                            PF3
                                                                            
                                                                            EMReadScreen panel_nav_check, 4, 2, 46
                                                                            If panel_nav_check <> "WRAP" Then
                                                                                PF3
                                                                                If activate_msg_boxes = True Then msgbox "Testing -- The script should now be at STAT/WRAP. If it is not, then stop here."
                                                                            End If

                                                                            If activate_msg_boxes = True Then msgbox "Testing -- No jobs panels existed. Created JOBS panel(s) through CM"
                                                                            
                                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No JOBS panels exist for household member number: " & employer_full_name & " that match the HIRE message. JOBS Panel and CASE/NOTE added for employer noted in HIRE message. Message should be deleted.")
                                                                            
                                                                        End If
                                                                    End If
                                                                End If
                                                            End if
                                                        End If
                                                    End If

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should not be deleted") Then
                                                        'The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should be deleted") Then
                                                        'There is a corresponding JOBS panel or a JOBS panel was created. The message can be deleted.
                                                        list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        dail_row = dail_row - 1
                                                    End If

                                                    'PF3 back to DAIL
                                                    PF3

                                                    If activate_msg_boxes = True Then MsgBox "The message has been processed and script will navigate back to DAIL now."

                                                    'Deactivate testing msgboxes as needed
                                                    ' activate_msg_boxes = False

                                                    'There is likely an issue with cases sometimes getting locked in background after adding new job and then navigating to next CSES message. Adding a background check here to avoid that issue.

                                                    'Attempt to get to STAT
                                                    EMWriteScreen "S", 6, 3
                                                    EMSendKey "<enter>"
                                                    EMReadScreen background_check, 25, 7, 30
                                                    If InStr(background_check, "A Background transaction") Then
                                                        EMWaitReady 2, 2000
                                                        Do
                                                            background_check = ""
                                                            PF3
                                                            EMWaitReady 2, 2000
                                                            EMWriteScreen "S", 6, 3
                                                            EMWaitReady 2, 2000
                                                            EMSendKey "<enter>"
                                                            EMWaitReady 2, 2000
                                                            EMReadScreen background_check, 25, 7, 30
                                                            If InStr(background_check, "A Background transaction") = 0 then Exit Do
                                                        Loop
                                                    End If
                                                    'Should be at STAT but need to double-check
                                                    EMReadScreen stat_summ_check, 4, 2, 46
                                                    EMReadScreen returned_to_SELF_check, 4, 2, 50
                                                    If stat_summ_check = "SUMM" Then 
                                                        'Successfully made it to STAT, can PF3 back to DAIL now
                                                        PF3
                                                    Else
                                                        msgbox "Delete after testing -- didn't make it back to STAT/SUMM. Check functionality"
                                                        If returned_to_SELF_check = "SELF" Then
                                                            'Script got bumped back to SELF, need to try to get back to DAIL while still acounting for background lock. Also need to reset the DAIL selection
                                                            Do
                                                                EMWriteScreen "DAIL", 16, 43
                                                                Call write_value_and_transmit("DAIL", 21, 70)
                                                                EMReadScreen SELF_check, 4, 2, 50
                                                                If SELF_check = "SELF" then
                                                                    PF3
                                                                    Pause 2
                                                                End if
                                                            Loop until SELF_check <> "SELF"

                                                            msgbox "3575 Delete after testing -- Escaped SELF loop, should be at DAIL/DAIL"

                                                            'Check to  make sure that we made it back to DAIL, it should maintain the case number
                                                            EMReadScreen back_to_dail_check, 8, 1, 72
                                                            If back_to_dail_check = "FMKDLAM6" Then
                                                                msgbox "3580 Delete after testing - should be back at DAIL/DAIL. Show back_to_dail_check >" & back_to_dail_check

                                                                'Navigate to CASE/CURR to force DAIL to reset and then PF3 back to get back to start of the DAIL
                                                                Call write_value_and_transmit("H", 6, 3)
                                                                PF3

                                                                msgbox "3586 Delete after testing -- Did it PF3 back to just DAIL?"

                                                                'Reset DAIL to only CSES messages
                                                                Call write_value_and_transmit("X", 4, 12)
                                                                EMWriteScreen "_", 7, 39
                                                                Call write_value_and_transmit("X", 10, 39)

                                                                'Script should now navigate to specific member name, or at least get close
                                                                EMWriteScreen case_name_to_return, 21, 25
                                                                transmit

                                                                msgbox "Delete after testing -- Made it back to DAIL. Should have reset the DAIL to CSES messages and got close to correct DAIL message"
                                                            Else

                                                                'Initial dialog - select whether to create a list or process a list
                                                                Dialog1 = ""
                                                                BeginDialog Dialog1, 0, 0, 306, 220, "Unable to return to DAIL. Double-check the issue."
                                                                
                                                                ButtonGroup ButtonPressed
                                                                    OkButton 205, 200, 40, 15
                                                                    CancelButton 245, 200, 40, 15
                                                                EndDialog

                                                                Do
                                                                    Dialog Dialog1
                                                                Loop until ButtonPressed = OK
                                                            End If
                                                        End If
                                                    End If

                                                ElseIf InStr(dail_msg, "REPORTED: CHILD REF NBR:") Then
                                                    'No action on these, simply note in spreadsheet that QI team to review

                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = "QI Review. CHILD NO LONGER RESIDES WITH CAREGIVER."

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                    'Update the excel spreadsheet with processing notes
                                                    objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                Else
                                                    msgbox "Testing -- A CSES message has appeared that does not meet either types - it will be SKIPPED. DAIL message is: " & dail_msg
                                                    
                                                    'No action on these, simply note in spreadsheet that QI team to review
                                                    
                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = "QI Review. DISB EXCESS CS (TYPE 43)."
                                                    
                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                    'Update the excel spreadsheet with processing notes
                                                    objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    
                                                    msgbox "Testing -- Ensure spreadsheet updated correctly with the CSES Type that cannot be processed"

                                                End If
                                            Else
                                                ' MsgBox "Hold in case need handling"
                                            End If

                                        End If

                                        'Increment the dail_excel_row so that data isn't overwritten
                                        dail_excel_row = dail_excel_row + 1
                                        
                                        'Increment dail_count for the dail array
                                        dail_count = dail_count + 1

                                        'In instances where the case details are not the final item in the array, need to exit the for loop
                                        Exit For

                                    End If 
                                Next

                            Else
                                'Add handling for messages that are not meeting any criteria. May not be necessary but have this just in case
                            End If
                                
                        End If
                    Else
                        'May not be needed but can add handling for cases that are not on valid case numbers list, just set processable to false and include processing note that it is likely out of county or privileged?
                    
                    End If
                            
                
                Else
                    'If dail_type is not CSES or HIRE then it is out of scope and there is no need to evaluate it

                End If

                ' Increment the stats counter
                stats_counter = stats_counter + 1
                
                dail_row = dail_row + 1

                'Checking for the last DAIL message. If it just processed the final message, the DAIL will appear blank but there is actually an invisible '_' at 6, 3. Handling to check for this and then navigate to the next page if needed. If it is on the last page, then it will exit the do loop 
                EMReadScreen next_dail_check, 7, dail_row, 3
                If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    'Attempt to navigate to the next page
                    PF8
                    EMReadScreen last_page_check, 21, 24, 2
                    'Check if the last page of the DAIL has been reached, also handles for situations where the last DAIL has been deleted and it displays a 'NO MESSAGES' warning
                    If last_page_check = "THIS IS THE LAST PAGE" or Instr(last_page_check, "NO MESSAGES") then
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

    'Update Stats Info

    'Calculate the manual time savings
    total_cases_evaluated = case_excel_row - 2
    STATS_manualtime = (total_cases_evaluated * 30) + (dail_msg_deleted_count * 90) + (not_processable_msg_count * 15) + (QI_flagged_msg_count * 30)

    'Activate the stats sheet
    objExcel.Worksheets("Stats").Activate
    objExcel.Cells(1, 2).Value = case_excel_row - 2
    objExcel.Cells(2, 2).Value = dail_excel_row - 2
    objExcel.Cells(3, 2).Value = not_processable_msg_count
    objExcel.Cells(4, 2).Value = dail_msg_deleted_count
    objExcel.Cells(5, 2).Value = QI_flagged_msg_count
    objExcel.Cells(6, 2).Value = timer - start_time ' script runtime
    objExcel.Cells(7, 2).Value = ((STATS_manualtime) - (timer - start_time)) / 60 'time savings from script


    'Finding the right folder to automatically save the file
    this_month = CM_mo & " " & CM_yr
    month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
    unclear_info_folder = replace(this_month, " ", "-") & " DAIL Unclear Info"
    report_date = replace(date, "/", "-")

    'saving the Excel file
    file_info = month_folder & "\" & unclear_info_folder & "\" & report_date & " Unclear Info" & " " & "CSES" & " " & dail_msg_deleted_count

    'Saves and closes the most recent Excel workbook with the Task based cases to process.
    objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"
    objExcel.ActiveWorkbook.Close
    objExcel.Application.Quit
    objExcel.Quit

    script_end_procedure_with_error_report("Success! Please review the list created for accuracy.")

End If

If HIRE_messages = 1 Then 

    'Create an array to track case details
    DIM HIRE_case_details_array()

    'constants for array
    const HIRE_case_maxis_case_number_const      = 0
    const HIRE_case_worker_const	             = 1
    const HIRE_active_programs_const             = 2
    const HIRE_pending_programs_const            = 3
    const HIRE_snap_status_const                 = 4
    const HIRE_snap_type_const                   = 5
    const HIRE_reporting_status_const            = 6
    const HIRE_sr_report_date_const              = 7 
    const HIRE_recertification_date_const        = 8
    const HIRE_MFIP_status_const                 = 9
    const HIRE_MFIP_MFSM_review_date_const       = 10
    const HIRE_MFIP_STAT_REVW_review_date_const  = 11
    const HIRE_GA_status_const                   = 12
    const HIRE_GA_reporting_status_const         = 13
    const HIRE_GA_budget_cycle_const             = 14
    const HIRE_GA_earned_income_const            = 15
    const HIRE_GA_GASM_review_date_const         = 16
    const HIRE_GA_STAT_REVW_review_date_const    = 17
    const HIRE_case_processing_notes_const       = 18
    const HIRE_processable_based_on_case_const   = 19
    const HIRE_case_excel_row_const              = 20

    'Opening the Excel file for list of DAIL messages
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True

    'Changes name of Excel sheet to DAIL Messages to capture details about in-scope DAIL messages
    ObjExcel.ActiveSheet.Name = "DAIL Messages"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "Case Number"
    objExcel.Cells(1, 2).Value = "X Number"
    objExcel.Cells(1, 3).Value = "DAIL Type"
    objExcel.Cells(1, 4).Value = "DAIL Month"
    objExcel.Cells(1, 5).Value = "DAIL Message"
    objExcel.Cells(1, 6).Value = "Full DAIL Message"
    objExcel.Cells(1, 7).Value = "Processing Notes for DAIL Message"

    FOR i = 1 to 7		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Creating second Excel sheet for compiling case details
    ObjExcel.Worksheets.Add().Name = "Case Details"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "Case Number"
    objExcel.Cells(1, 2).Value = "X Number"
    objExcel.Cells(1, 3).Value = "Active Programs"
    objExcel.Cells(1, 4).Value = "Pending Programs"
    objExcel.Cells(1, 5).Value = "SNAP Status"
    objExcel.Cells(1, 6).Value = "SNAP Type"
    objExcel.Cells(1, 7).Value = "SNAP Reporting Status"
    objExcel.Cells(1, 8).Value = "SNAP SR Report Date"
    objExcel.Cells(1, 9).Value = "SNAP Recertification Date"
    objExcel.Cells(1, 10).Value = "MFIP Status"
    objExcel.Cells(1, 11).Value = "MFIP MFSM Review Date"
    objExcel.Cells(1, 12).Value = "MFIP STAT/REVW Review Date"
    objExcel.Cells(1, 13).Value = "GA Status"
    objExcel.Cells(1, 14).Value = "GA Reporting Status"
    objExcel.Cells(1, 15).Value = "GA Budget Cycle"
    objExcel.Cells(1, 16).Value = "GA Earned Income"
    objExcel.Cells(1, 17).Value = "GA GASM Review Date"
    objExcel.Cells(1, 18).Value = "GA STAT/REVW Review Date"
    objExcel.Cells(1, 19).Value = "Processing Notes for Case"
    objExcel.Cells(1, 20).Value = "Processable based on Case Details"

    FOR i = 1 to 20		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Creates sheet to track stats for the script
    ObjExcel.Worksheets.Add().Name = "Stats"

    'Setting counters at 0
    STATS_counter = STATS_counter - 1
    not_processable_msg_count = 0
    dail_msg_deleted_count = 0
    QI_flagged_msg_count = 0

    'Enters info about runtime for the benefit of folks using the script
    objExcel.Cells(1, 1).Value = "Cases Evaluated:"
    objExcel.Cells(2, 1).Value = "Evaluated DAIL Messages:"
    objExcel.Cells(3, 1).Value = "Unprocessable DAIL Messages:"
    objExcel.Cells(4, 1).Value = "Deleted DAIL Messages:"
    objExcel.Cells(5, 1).Value = "QI Flagged Messages:"
    objExcel.Cells(6, 1).Value = "Script run time (in seconds):"
    objExcel.Cells(7, 1).Value = "Estimated time savings by using script (in minutes):"

    FOR i = 1 to 7		'formatting the cells'
        objExcel.Cells(i, 1).Font.Bold = True		'bold font'
        ObjExcel.rows(i).NumberFormat = "@" 		'formatting as text
        objExcel.columns(1).AutoFit()				'sizing the columns'
    NEXT

    'Add details for tracking TIKLs
    'Creating second Excel sheet for compiling case details
    ObjExcel.Worksheets.Add().Name = "HIRE TIKLs"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "Case Number"
    objExcel.Cells(1, 2).Value = "Case Name"
    objExcel.Cells(1, 3).Value = "DAIL Type"
    objExcel.Cells(1, 4).Value = "TIKL Date"
    objExcel.Cells(1, 5).Value = "TIKL Message"
    objExcel.Cells(1, 6).Value = "Action Taken on TIKL"

    FOR i = 1 to 6		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Setting starting row for TIKLs sheet
    TIKL_excel_row = 2

    ReDim DAIL_message_array(7, 0)
    'Incrementor for the array
    Dail_count = 0

    'Sets variable for the Excel row to export data to Excel sheet
    dail_excel_row = 2

    ReDim HIRE_case_details_array(HIRE_case_excel_row_const, 0)

    'Incrementor for the array
    case_count = 0

    'Sets variable for the Excel row to export data to Excel sheet
    case_excel_row = 2

    'Reset the array 
    ReDim PMI_and_ref_nbr_array(3, 0)

    'Incrementor for the array
    member_count = 0

    For each worker in worker_array

        'Clearing out MAXIS case number so that it doesn't carry forward from previous case
        MAXIS_case_number = ""
        
        'Resetting all of the string lists
        'Creating initial string for tracking list of valid case numbers pulled from REPT/ACTV. This is used to avoid triggering a privileged case and losing connection to DAIL
        valid_case_numbers_list = "*"

        'Create list of case numbers to be used for comparison purposes as the script navigates through the DAIL
        list_of_all_case_numbers = "*"

        'Create list of DAIL messages that should be deleted for NDNH messages where the employment is known. If a DAIL message matches, then it will be deleted. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc.
        list_of_DAIL_messages_to_delete_NDNH_known = "*"

        'Create list of DAIL messages that should be deleted for NDNH messages where the employment is NOT known. If a DAIL message matches, then it will be deleted. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc.
        list_of_DAIL_messages_to_delete_NDNH_not_known = "*"

        'Create list of DAIL messages that should be deleted for SDNH messages since these can just be deleted. If a DAIL message matches, then it will be deleted. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc.
        list_of_DAIL_messages_to_delete_SDNH = "*"

        'Create list of DAIL messages that should be skipped. If a DAIL message matches, then the script will skip past it to next DAIL row. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc. 
        If list_of_DAIL_messages_to_skip = "" then list_of_DAIL_messages_to_skip = "*"

        'Create strings for tracking NDNH messages
        list_of_NDNH_messages_standard_format = "*"

        'Create strings for tracking TIKLs to be deleted
        list_of_TIKLs_to_delete = "*"

        'Formatting the worker so there are no errors
        worker = trim(ucase(worker))

        'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason					
        back_to_self	

        Call navigate_to_MAXIS_screen("REPT", "ACTV")
        EMWriteScreen worker, 21, 13
        TRANSMIT
        EMReadScreen user_worker, 7, 21, 71
        EMReadScreen p_worker, 7, 21, 13
        IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

        IF worker_number = "X127CCL" or worker = "127CCL" THEN
            DO
                EmReadScreen worker_confirmation, 20, 3, 11 'looking for CENTURY PLAZA CLOSED
                EMWaitReady 0, 0
            LOOP UNTIL worker_confirmation = "CENTURY PLAZA CLOSED"
        END IF

        'Skips workers with no info
        EMReadScreen has_content_check, 1, 7, 8
        If has_content_check <> " " then
            'Grabbing each case number on screen
            Do
                'Set variable for next do...loop
                MAXIS_row = 7
                'Checking for the last page of cases.
                EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
                EMReadscreen number_of_pages, 4, 3, 76 'getting page number because to ensure it doesnt fail'
                number_of_pages = trim(number_of_pages)
                Do
                    EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number

                    'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
                    MAXIS_case_number = trim(MAXIS_case_number)
                    If MAXIS_case_number <> "" and instr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") <> 0 then exit do
                    valid_case_numbers_list = trim(valid_case_numbers_list & MAXIS_case_number & "*")

                    If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

                    MAXIS_row = MAXIS_row + 1
                    MAXIS_case_number = ""			'Blanking out variable
                Loop until MAXIS_row = 19
                PF8
            Loop until last_page_check = "THIS IS THE LAST PAGE"
        END IF

        'Navigates to DAIL to pull DAIL messages
        MAXIS_case_number = ""
        CALL navigate_to_MAXIS_screen("DAIL", "PICK")
        EMWriteScreen "_", 7, 39    'blank out ALL selection
        'Selects INFO (HIRE) DAIL Type based on dialog selection
        EMWriteScreen "X", 13, 39
        transmit

        'Enter the worker number on DAIL to pull up DAIL messages
        Call write_value_and_transmit(worker, 21, 6)
        'Transmits past not your dail message
        transmit 

        'Reads where the count of DAILs is listed. Used to verify DAIL is not empty.
        EMReadScreen number_of_dails, 1, 3, 67		

        DO
        'If this space is blank the rest of the DAIL reading is skipped
            If number_of_dails = " " Then 
                exit do		
            End if
            'Because the script brings each new case to the top of the page, dail_row starts at 6.
            dail_row = 6	

            DO
                dail_type = ""
                dail_msg = ""
                dail_month = ""
                MAXIS_case_number = ""
                actionable_dail = ""
                renewal_6_month_check = ""

                'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                EMReadScreen new_case, 8, dail_row, 63
                new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN 
                    'If there is NOT a new case number, the script will top the message
                    Call write_value_and_transmit("T", dail_row, 3)
                ELSEIF new_case = "CASE NBR" THEN
                    'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                End if

                'Resets the DAIL row since the message has now been topped
                dail_row = 6  

                'Determines the DAIL Type
                EMReadScreen dail_type, 4, dail_row, 6
                dail_type = trim(dail_type)

                'Reads the DAIL msg to determine if it is an out-of-scope message
                EMReadScreen dail_msg, 61, dail_row, 20
                dail_msg = trim(dail_msg)

                'List of out of scope messages pulled from non-actionable dails function
                If instr(dail_msg, "AMT CHILD SUPP MOD/ORD") OR _
                    instr(dail_msg, "AP OF CHILD REF NBR:") OR _
                    instr(dail_msg, "ADDRESS DIFFERS W/ CS RECORDS:") OR _
                    instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN LBUD IN THE MONTH") OR _
                    instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN SBUD IN THE MONTH") OR _
                    instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                    instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                    instr(dail_msg, "CHILD SUPP PAYMTS PD THRU THE COURT/AGENCY FOR CHILD") OR _
                    instr(dail_msg, "COMPLETE INFC PANEL") OR _
                    instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
                    instr(dail_msg, "IS COOPERATING WITH CHILD SUPPORT") OR _
                    instr(dail_msg, "IS NOT COOPERATING WITH CHILD SUPPORT") OR _
                    instr(dail_msg, "NAME DIFFERS W/ CS RECORDS:") OR _
                    instr(dail_msg, "PATERNITY ON CHILD REF NBR") OR _
                    instr(dail_msg, "REPORTED NAME CHG TO:") OR _
                    instr(dail_msg, "BENEFITS RETURNED, IF IOC HAS NEW ADDRESS") OR _
                    instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") OR _
                    instr(dail_msg, "CASE NOT AUTO-APPROVED - HRF/SR/RECERT DUE") OR _
                    instr(dail_msg, "CHANGE IN BUDGET CYCLE") OR _
                    instr(dail_msg, "COMPLETE ELIG IN FIAT") OR _
                    instr(dail_msg, "COUNTED IN LBUD AS UNEARNED INCOME") OR _
                    instr(dail_msg, "COUNTED IN SBUD AS UNEARNED INCOME") OR _
                    instr(dail_msg, "HAS EARNED INCOME IN 6 MONTH BUDGET BUT NO DCEX PANEL") OR _
                    instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") OR _
                    instr(dail_msg, "NEW ELIG RESULTS EXIST") OR _
                    instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") OR _
                    instr(dail_msg, "WARNING MESSAGES EXIST") OR _
                    instr(dail_msg, "ADDR CHG*CHK SHEL") OR _
                    instr(dail_msg, "APPLCT ID CHNGD") OR _
                    instr(dail_msg, "CASE AUTOMATICALLY DENIED") OR _
                    instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") OR _
                    instr(dail_msg, "CASE NOTE ENTERED BY") OR _
                    instr(dail_msg, "CASE NOTE TRANSFER FROM") OR _
                    instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") OR _
                    instr(dail_msg, "CASE XFER") OR _
                    instr(dail_msg, "CHANGE REPORT FORM SENT ON") OR _
                    instr(dail_msg, "DIRECT DEPOSIT STATUS") OR _
                    instr(dail_msg, "EFUNDS HAS NOTIFIED DHS THAT THIS CLIENT'S EBT CARD") OR _
                    instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") OR _
                    instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") OR _
                    instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") OR _
                    instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") OR _
                    instr(dail_msg, "NOT ACCESSED FOR 300 DAYS,SPEC NOT") OR _
                    instr(dail_msg, "PMI MERGED") OR _
                    instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") OR _
                    instr(dail_msg, "THIS CASE IS ERROR PRONE") OR _
                    instr(dail_msg, "EMPL SERV REF DATE IS > 60 DAYS; CHECK ES PROVIDER RESPONSE") OR _
                    instr(dail_msg, "LAST GRADE COMPLETED") OR _
                    instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
                    instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
                    instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") OR _
                    instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") OR _
                    instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE") OR _
                    instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") OR _
                    instr(dail_msg, "GA: STATUS IS PENDING - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "GA: STATUS IS REIN OR SUSPEND - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") or _
                    instr(dail_msg, "GRH: APPROVED VERSION EXISTS FOR JANUARY - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS") OR _
                    instr(dail_msg, "MSA RECERT DUE - NOT AUTO") or _
                    instr(dail_msg, "MSA IN PENDING STATUS - NOT AUTO") or _
                    instr(dail_msg, "APPROVED MSA VERSION EXISTS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") or _
                    instr(dail_msg, "GRH: STATUS IS REIN, PENDING OR SUSPEND - NOT AUTO") OR _
                    instr(dail_msg, "SDNH NEW JOB DETAILS FOR MEMB 00") OR _
                    instr(dail_msg, "SNAP: PENDING OR STAT EDITS EXIST") OR _
                    instr(dail_msg, "SNAP: REIN STATUS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: APPROVED VERSION ALREADY EXISTS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: AUTO-APPROVED - PREVIOUS UNAPPROVED VERSION EXISTS") OR _
                    instr(dail_msg, "SSN DIFFERS W/ CS RECORDS") OR _
                    instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                    instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED CASE WITH SANCTION") OR _
                    instr(dail_msg, "DWP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                    instr(dail_msg, "IV-D NAME DISCREPANCY") OR _
                    instr(dail_msg, "CHECK HAS BEEN APPROVED") OR _
                    instr(dail_msg, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                    instr(dail_msg, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                    instr(dail_msg, "- TRANS #") OR _
                    instr(dail_msg, "RSDI UPDATED - (REF") OR _
                    instr(dail_msg, "SSI UPDATED - (REF") OR _
                    instr(dail_msg, "SNAP ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS") then 
                        actionable_dail = False
                Else
                    actionable_dail = True
                End If

                If actionable_dail = True and dail_type = "HIRE" Then
                    'Script compiles a list of all of the NDNH, but only for active cases that are not privileged or out of county

                    'Read the MAXIS Case Number, if it is a new case number then pull case details. If it is not a new case number, then do not pull new case details.
                    
                    EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                    MAXIS_case_number = trim(MAXIS_case_number)

                    If InStr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") Then
                        'If the case is in the valid_case_numbers_list, then it can be evaluated. If it is NOT in the valid_case_numbers_list then it is likely privileged or out of county so it will be skipped

                        If InStr(dail_msg, "NDNH MEMB") Then
                            ' Script reads the full DAIL message for NDNH messages and adds to a string to compare SDNH messages against when it runs through the X number again to actually process the messages

                            'Open the full HIRE message
                            Call write_value_and_transmit("X", dail_row, 3)

                            'Delete after testing - trying to figure out when and why script sometimes does not clear the X
                            EmReadScreen multiple_selections_error_check, 20, 24, 2
                            If InStr(multiple_selections_error_check, "YOU MAY ONLY SELECT") Then msgbox "4126 It failed to clear the previous X" 

                            'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                            'Set row and col
                            row = 1
                            col = 1
                            EMSearch "Case Number: ", row, col
                            EMReadScreen HIRE_case_number, 10, row, col + 13
                            HIRE_case_number = trim(HIRE_case_number)

                            row = 1
                            col = 1
                            EMSearch "Case Name: ", row, col
                            EMReadScreen HIRE_case_name, 25, row, col + 11
                            HIRE_case_name = trim(HIRE_case_name)

                            row = 1
                            col = 1
                            EMSearch "NDNH MEMB ", row, col
                            EMReadScreen HIRE_memb_number, 2, row, col + 10
                            HIRE_memb_number = trim(HIRE_memb_number)

                            row = 1
                            col = 1
                            EMSearch "DATE HIRED   :", row, col
                            EMReadScreen date_hired, 10, row, col + 15
                            date_hired = trim(date_hired)

                            row = 1
                            col = 1
                            EMSearch "EMPLOYER: ", row, col
                            EMReadScreen HIRE_employer_name, 20, row, col + 10
                            HIRE_employer_name = trim(HIRE_employer_name)

                            row = 1
                            col = 1
                            EMSearch "MAXIS NAME   :", row, col
                            EMReadScreen HIRE_maxis_name, 57, row, col + 15
                            HIRE_maxis_name = trim(HIRE_maxis_name)

                            row = 1
                            col = 1
                            EMSearch "NEW HIRE NAME:", row, col
                            EMReadScreen HIRE_new_hire_name, 57, row, col + 15
                            HIRE_new_hire_name = trim(HIRE_new_hire_name)

                            'Standard NDNH format is *[Case Number]-[Case Name]-[Memb ##]-[Date Hired with slashes]-[Employer - first 20 characters]-[Maxis name]-[new hire name]*
                            hire_ndnh_message_standardized = HIRE_case_number & "-" & HIRE_case_name & "-" & HIRE_memb_number & "-" & date_hired & "-" & HIRE_employer_name & "-" & HIRE_maxis_name & "-" & HIRE_new_hire_name
                            list_of_NDNH_messages_standard_format = list_of_NDNH_messages_standard_format & hire_ndnh_message_standardized & "*"  
                            'Transmit back to DAIL
                            transmit
                        End If
                    End If
                End If
                            
                dail_row = dail_row + 1

                EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
                If message_error = "NO MESSAGES" then exit do

                'Checking for the last DAIL message. If it just processed the final message, the DAIL will appear blank but there is actually an invisible '_' at 6, 3. Handling to check for this and then navigate to the next page if needed. If it is on the last page, then it will exit the do loop 
                EMReadScreen next_dail_check, 7, dail_row, 3
                If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    'Attempt to navigate to the next page
                    PF8
                    EMReadScreen last_page_check, 21, 24, 2
                    'Check if the last page of the DAIL has been reached, also handles for situations where the last DAIL has been deleted and it displays a 'NO MESSAGES' warning
                    If last_page_check = "THIS IS THE LAST PAGE" or Instr(last_page_check, "NO MESSAGES") then
                        all_done = true
                        exit do
                    Else
                        dail_row = 6
                    End if
                End if
            LOOP
            IF all_done = true THEN exit do
        LOOP

        'Now that the script has compiled a string of all of the NDNH messages, it will now evaluate the individual messages to determine if there is a duplicate SDNH, or if it can process the SDNH or NDNH message
        'Reset the all_done so that it doesn't exit after the first run unintentionally
        all_done = ""

        'Navigates to DAIL to pull DAIL messages and start at beginning again
        'Go back to start (" A" used to get as close to first case as possible)
        loop_count = 0
        EMReadScreen number_of_dails, 1, 3, 67	
        If number_of_dails <> " " Then 
            Call write_value_and_transmit(" A", 21, 25)

            Do
                PF7
                EMReadScreen first_page_check, 37, 24, 2
                If first_page_check = "YOU MAY ONLY SCROLL FORWARD FROM HERE" Then Exit Do
                EMReadScreen number_of_dails, 1, 3, 67
                If number_of_dails = " " Then 
                    exit do
                End If
                loop_count = loop_count + 1
                If loop_count = 5 then MsgBox "Testing -- it is stuck in a loop"
            Loop
        End If

        'Reads where the count of DAILs is listed. Used to verify DAIL is not empty.
        EMReadScreen number_of_dails, 1, 3, 67		

        DO
            'If this space is blank the rest of the DAIL reading is skipped
            If number_of_dails = " " Then exit do		
            'Because the script brings each new case to the top of the page, dail_row starts at 6.
            dail_row = 6	

            DO
                dail_type = ""
                dail_msg = ""
                dail_month = ""
                MAXIS_case_number = ""
                actionable_dail = ""
                renewal_6_month_check = ""
                SNAP_active = ""
                MFIP_active = ""
                GA_Active = ""
                SNAP_MFIP_GA_active = ""
                Other_programs_active_or_pending = ""

                'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                EMReadScreen new_case, 8, dail_row, 63
                new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN 
                    'If there is NOT a new case number, the script will top the message
                    Call write_value_and_transmit("T", dail_row, 3)
                ELSEIF new_case = "CASE NBR" THEN
                    'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                End if

                'Resets the DAIL row since the message has now been topped
                dail_row = 6  

                'Determines the DAIL Type
                EMReadScreen dail_type, 4, dail_row, 6
                dail_type = trim(dail_type)

                'Reads the DAIL msg to determine if it is an out-of-scope message
                EMReadScreen dail_msg, 61, dail_row, 20
                dail_msg = trim(dail_msg)

                'List of out of scope messages pulled from non-actionable dails function
                If instr(dail_msg, "AMT CHILD SUPP MOD/ORD") OR _
                    instr(dail_msg, "AP OF CHILD REF NBR:") OR _
                    instr(dail_msg, "ADDRESS DIFFERS W/ CS RECORDS:") OR _
                    instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN LBUD IN THE MONTH") OR _
                    instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN SBUD IN THE MONTH") OR _
                    instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                    instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                    instr(dail_msg, "CHILD SUPP PAYMTS PD THRU THE COURT/AGENCY FOR CHILD") OR _
                    instr(dail_msg, "COMPLETE INFC PANEL") OR _
                    instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
                    instr(dail_msg, "IS COOPERATING WITH CHILD SUPPORT") OR _
                    instr(dail_msg, "IS NOT COOPERATING WITH CHILD SUPPORT") OR _
                    instr(dail_msg, "NAME DIFFERS W/ CS RECORDS:") OR _
                    instr(dail_msg, "PATERNITY ON CHILD REF NBR") OR _
                    instr(dail_msg, "REPORTED NAME CHG TO:") OR _
                    instr(dail_msg, "BENEFITS RETURNED, IF IOC HAS NEW ADDRESS") OR _
                    instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") OR _
                    instr(dail_msg, "CASE NOT AUTO-APPROVED - HRF/SR/RECERT DUE") OR _
                    instr(dail_msg, "CHANGE IN BUDGET CYCLE") OR _
                    instr(dail_msg, "COMPLETE ELIG IN FIAT") OR _
                    instr(dail_msg, "COUNTED IN LBUD AS UNEARNED INCOME") OR _
                    instr(dail_msg, "COUNTED IN SBUD AS UNEARNED INCOME") OR _
                    instr(dail_msg, "HAS EARNED INCOME IN 6 MONTH BUDGET BUT NO DCEX PANEL") OR _
                    instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") OR _
                    instr(dail_msg, "NEW ELIG RESULTS EXIST") OR _
                    instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") OR _
                    instr(dail_msg, "WARNING MESSAGES EXIST") OR _
                    instr(dail_msg, "ADDR CHG*CHK SHEL") OR _
                    instr(dail_msg, "APPLCT ID CHNGD") OR _
                    instr(dail_msg, "CASE AUTOMATICALLY DENIED") OR _
                    instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") OR _
                    instr(dail_msg, "CASE NOTE ENTERED BY") OR _
                    instr(dail_msg, "CASE NOTE TRANSFER FROM") OR _
                    instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") OR _
                    instr(dail_msg, "CASE XFER") OR _
                    instr(dail_msg, "CHANGE REPORT FORM SENT ON") OR _
                    instr(dail_msg, "DIRECT DEPOSIT STATUS") OR _
                    instr(dail_msg, "EFUNDS HAS NOTIFIED DHS THAT THIS CLIENT'S EBT CARD") OR _
                    instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") OR _
                    instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") OR _
                    instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") OR _
                    instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") OR _
                    instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") OR _
                    instr(dail_msg, "NOT ACCESSED FOR 300 DAYS,SPEC NOT") OR _
                    instr(dail_msg, "PMI MERGED") OR _
                    instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") OR _
                    instr(dail_msg, "THIS CASE IS ERROR PRONE") OR _
                    instr(dail_msg, "EMPL SERV REF DATE IS > 60 DAYS; CHECK ES PROVIDER RESPONSE") OR _
                    instr(dail_msg, "LAST GRADE COMPLETED") OR _
                    instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
                    instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
                    instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") OR _
                    instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") OR _
                    instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE") OR _
                    instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") OR _
                    instr(dail_msg, "GA: STATUS IS PENDING - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "GA: STATUS IS REIN OR SUSPEND - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") or _
                    instr(dail_msg, "GRH: APPROVED VERSION EXISTS FOR JANUARY - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS") OR _
                    instr(dail_msg, "MSA RECERT DUE - NOT AUTO") or _
                    instr(dail_msg, "MSA IN PENDING STATUS - NOT AUTO") or _
                    instr(dail_msg, "APPROVED MSA VERSION EXISTS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") or _
                    instr(dail_msg, "GRH: STATUS IS REIN, PENDING OR SUSPEND - NOT AUTO") OR _
                    instr(dail_msg, "SDNH NEW JOB DETAILS FOR MEMB 00") OR _
                    instr(dail_msg, "SNAP: PENDING OR STAT EDITS EXIST") OR _
                    instr(dail_msg, "SNAP: REIN STATUS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: APPROVED VERSION ALREADY EXISTS - NOT AUTO-APPROVED") OR _
                    instr(dail_msg, "SNAP: AUTO-APPROVED - PREVIOUS UNAPPROVED VERSION EXISTS") OR _
                    instr(dail_msg, "SSN DIFFERS W/ CS RECORDS") OR _
                    instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                    instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED CASE WITH SANCTION") OR _
                    instr(dail_msg, "DWP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                    instr(dail_msg, "IV-D NAME DISCREPANCY") OR _
                    instr(dail_msg, "CHECK HAS BEEN APPROVED") OR _
                    instr(dail_msg, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                    instr(dail_msg, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                    instr(dail_msg, "- TRANS #") OR _
                    instr(dail_msg, "RSDI UPDATED - (REF") OR _
                    instr(dail_msg, "SSI UPDATED - (REF") OR _
                    instr(dail_msg, "SNAP ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS") then 
                        actionable_dail = False
                Else
                    actionable_dail = True
                End If

                If actionable_dail = True AND dail_type = "HIRE" Then
                    'Read the MAXIS Case Number, if it is a new case number then pull case details. If it is not a new case number, then do not pull new case details.
                    
                    EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                    MAXIS_case_number = trim(MAXIS_case_number)

                    If InStr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") Then
                        'If the case is in the valid_case_numbers_list, then it can be evaluated. If it is NOT in the valid_case_numbers_list then it is likely privileged or out of county so it will be skipped

                        If Instr(list_of_all_case_numbers, "*" & MAXIS_case_number & "*") = 0 Then
                            'If the MAXIS case number is NOT in the list of all case numbers, then it is a new case number and the script will gather case details
                            
                            'Redim the case details array and add to array
                            ReDim Preserve HIRE_case_details_array(HIRE_case_excel_row_const, case_count)
                            HIRE_case_details_array(HIRE_case_maxis_case_number_const, case_count) = MAXIS_case_number
                            HIRE_case_details_array(HIRE_case_worker_const, case_count) = worker
                    
                            'Since case number is not in list of all case numbers, add it to the list
                            list_of_all_case_numbers = list_of_all_case_numbers & MAXIS_case_number & "*"

                            'Navigate to CASE/CURR to pull case details 
                            Call write_value_and_transmit("H", dail_row, 3)

                            'Handling if the case is out of county
                            EmReadscreen worker_county, 4, 21, 14
                            If worker_county <> worker_county_code then
                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = "Out-of-County Case"
                                HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                            Else
                                'Pull case details from CASE/CURR, maintains connection to DAIL
                                Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

                                'Split list of active programs into an array to validate
                                If trim(list_active_programs) <> "" Then 
                                    split_list_active_programs = split(list_active_programs, ", ")

                                    i = 0
                                    Do
                                        If split_list_active_programs(i) = "SNAP" Then 
                                            SNAP_active = True
                                            SNAP_MFIP_GA_active = True
                                        ElseIf split_list_active_programs(i) = "MFIP" Then 
                                            MFIP_active = True
                                            SNAP_MFIP_GA_active = True
                                        ElseIf split_list_active_programs(i) = "GA" Then 
                                            GA_active = True
                                            SNAP_MFIP_GA_active = True
                                        Else
                                            'If it is a program other than SNAP, GA, and/or MFIP then we will need to skip this case
                                            other_programs_active_or_pending = other_programs_active_or_pending & split_list_active_programs(i) & ", "
                                        End If
                                        i = i + 1
                                    Loop until i = ubound(split_list_active_programs) + 1
                                End If

                                If list_pending_programs <> "" then other_programs_active_or_pending = other_programs_active_or_pending & list_pending_programs

                                If activate_msg_boxes = True Then msgbox "Delete after Testing -- SNAP_active = " & snap_active & vbcr & vbcr & "MFIP_active = " & MFIP_active & vbcr & vbcr & "GA_active = " & GA_active & vbcr & vbcr & "other_programs_active_or_pending = " & other_programs_active_or_pending & vbcr & vbcr & "case_active = " & case_active

                                'Update array with active and pending programs, and SNAP and MFIP statuses
                                HIRE_case_details_array(HIRE_active_programs_const, case_count) = list_active_programs
                                HIRE_case_details_array(HIRE_pending_programs_const, case_count) = list_pending_programs
                                HIRE_case_details_array(HIRE_SNAP_status_const, case_count) = trim(SNAP_status)
                                HIRE_case_details_array(HIRE_MFIP_status_const, case_count) = trim(MFIP_status)
                                HIRE_case_details_array(HIRE_GA_status_const, case_count) = trim(GA_status)

                                'Function (determine_program_and_case_status_from_CASE_CURR) sets dail_row equal to 4 so need to reset it.
                                dail_row = 6

                                If case_active = TRUE and SNAP_MFIP_GA_active = True and other_programs_active_or_pending = "" Then
                                    If SNAP_active = True Then
                                        'The case is active on SNAP, will gather more details about SNAP

                                        'Ensure that we are viewing ELIG/FS for the current month, not the dail message month
                                        EMWriteScreen MAXIS_footer_month, 20, 54
                                        EMWriteScreen MAXIS_footer_year, 20, 57

                                        'Navigate to ELIG/FS from CASE/CURR to maintain tie to DAIL
                                        EMWriteScreen "ELIG", 20, 22
                                        Call write_value_and_transmit("FS  ", 20, 69)

                                        EMReadScreen no_SNAP, 10, 24, 2
                                        If no_SNAP = "NO VERSION" then						'NO SNAP version means no determination
                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            Else
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            End If
                                            HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
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
                                                If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                Else
                                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                End If
                                                HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                            Elseif app_status = "APPROVED" then
                                                EMReadScreen vers_number, 1, status_row, 23
                                                Call write_value_and_transmit(vers_number, 18, 54)
                                                Call write_value_and_transmit("FSSM", 19, 70)

                                                EmReadscreen reporting_status, 12, 8, 31
                                                reporting_status = trim(reporting_status)

                                                'Read for UHFS
                                                EmReadscreen UHFS_status_check, 16, 4, 3
                                                If UHFS_status_check = "'UNCLE HARRY' FS" Then 
                                                    HIRE_case_details_array(HIRE_snap_type_const, case_count) = "UHFS"
                                                Else
                                                    HIRE_case_details_array(HIRE_snap_type_const, case_count) = "SNAP"
                                                End If
                                                
                                                If reporting_status = "SIX MONTH" Then
                                                    'Navigate to STAT/REVW to confirm recertification and SR report date
                                                    EMWriteScreen "STAT", 19, 22
                                                    EMWaitReady 0, 0
                                                    Call write_value_and_transmit("REVW", 19, 70)
                                                    
                                                    EMWaitReady 0, 0
                                                    EmReadscreen error_prone_check, 6, 2, 51

                                                    If InStr(error_prone_check, "ERRR") Then
                                                        transmit
                                                        EMWaitReady 0, 0
                                                    End If

                                                    'Pause here as it sometimes errors
                                                    EMWaitReady 0, 0
                                                    'Open the FS screen
                                                    EMWriteScreen "X", 5, 58
                                                    EMWaitReady 0, 0
                                                    Transmit
                                                    EMWaitReady 0, 0

                                                    EMReadScreen food_support_reports_check, 20, 5, 30
                                                    If food_support_reports_check <> "Food Support Reports" Then 
                                                        'Pause here as it sometimes errors
                                                        EMWaitReady 0, 0
                                                        'Open the FS screen
                                                        EMWriteScreen "X", 5, 58
                                                        EMWaitReady 0, 0
                                                        Transmit
                                                        EMWaitReady 0, 0
                                                        EMReadScreen food_support_reports_check, 20, 5, 30
                                                        If food_support_reports_check <> "Food Support Reports" Then MsgBox "Testing -- FS Screen attempt 2 did not work. Try rerunning script again."
                                                    End If

                                                    EmReadscreen sr_report_date, 8, 9, 26
                                                    EmReadscreen recertification_date, 8, 9, 64

                                                    'Add handling for missing SR report date or recertification
                                                    'Adds slashes to dates then converts to datedate from string to date
                                                    If sr_report_date = "__ 01 __" Then
                                                        sr_report_date = "SR Report Date is Missing"
                                                    Else
                                                        sr_report_date = replace(sr_report_date, " ", "/")
                                                        sr_report_date = DateAdd("m", 0, sr_report_date)
                                                    End If

                                                    If recertification_date = "__ 01 __" Then
                                                        recertification_date = "Recertification Date is Missing"
                                                    Else
                                                        recertification_date = replace(recertification_date, " ", "/")
                                                        recertification_date = DateAdd("m", 0, recertification_date)
                                                    End If
                            
                                                    If sr_report_date <> "SR Report Date is Missing" and recertification_date <> "Recertification Date is Missing" Then 
                                                        renewal_6_month_difference = DateDiff("M", sr_report_date, recertification_date)

                                                        If renewal_6_month_difference = "6" or renewal_6_month_difference = "-6" then 
                                                            renewal_6_month_check = True
                                                        Else 
                                                            renewal_6_month_check = False

                                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; SR Report Date and Recertification are not 6 months apart"
                                                            Else
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "SR Report Date and Recertification are not 6 months apart"
                                                            End If
                                                        End if

                                                        If DateDiff("m", footer_month_day_year, sr_report_date) < 0 AND DateDiff("m", footer_month_day_year, recertification_date) < 0 Then
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Both review dates are before CM and has not been updated correctly 4628"
                                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; SNAP Review Dates are prior to current month. Case should be reviewed."
                                                            Else
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "SNAP Review Dates are prior to current month. Case should be reviewed."
                                                            End If
                                                            HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                        End If
                                                    Else
                                                        renewal_6_month_check = False
                                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; SR Report Date and/or Recertification Date is missing"
                                                        Else
                                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "SR Report Date and/or Recertification Date is missing"
                                                        End If
                                                    End If
                                                    
                                                    'Close the FS screen
                                                    transmit
                                                Else
                                                    sr_report_date = "N/A"
                                                    recertification_date = "N/A"

                                                End If

                                            End If
                                            
                                            'Update the array with new case details
                                            HIRE_case_details_array(HIRE_reporting_status_const, case_count) = reporting_status
                                            HIRE_case_details_array(HIRE_recertification_date_const, case_count) = trim(recertification_date)
                                            HIRE_case_details_array(HIRE_sr_report_date_const, case_count) = trim(sr_report_date)

                                        End If
                                    End If

                                    If MFIP_active = True Then
                                        'Navigate to MFSM panel to confirm review date
                                        'Navigate to STAT/REVW to confirm review date

                                        'To ensure starting from DAIL, PF3 to get back to DAIL then navigate back to CASE/CURR
                                        'Back to DAIL
                                        PF3
                                        'Navigate back to CASE/CURR
                                        Call write_value_and_transmit("H", dail_row, 3)
                                        'Update the footer month/year and then navigate to ELIG/GA

                                        'Ensure that we are viewing ELIG/FS for the current month, not the dail message month
                                        EMWriteScreen MAXIS_footer_month, 20, 54
                                        EMWriteScreen MAXIS_footer_year, 20, 57
                                        
                                        'Navigate to ELIG/GA from CASE/CURR
                                        EMWriteScreen "ELIG", 20, 22
                                        Call write_value_and_transmit("MFIP", 20, 69)

                                        EMReadScreen no_MFIP, 10, 24, 2
                                        If no_MFIP = "NO VERSION" then						'NO GA version means no determination
                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; No version of MFIP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            Else
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "No version of MFIP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            End If
                                            HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                        Else
                                            EMWriteScreen "99", 20, 79
                                            transmit
                                            'This brings up the GA versions of eligibility results to search for approved versions
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
                                                If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; No approved eligibility results for MFIP exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                Else
                                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "No approved eligibility results for MFIP exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                End If
                                                HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                            Elseif app_status = "APPROVED" then
                                                'Check for earned income
                                                EMReadScreen vers_number, 1, status_row, 23
                                                Call write_value_and_transmit(vers_number, 18, 54)
                                                'Navigate to MFSM panel to verify earned income total
                                                Call write_value_and_transmit("MFSM", 20, 71)
                                                EmReadScreen MFSM_panel_check, 4, 3, 47
                                                If MFSM_panel_check <> "MFSM" Then msgbox "Testing -- 4561 Error unable to reach MFSM"
                                                
                                                'Read eligibility review date from MFSM panel
                                                EMReadScreen MFIP_MFSM_review_date, 8, 11, 31
                                                HIRE_case_details_array(HIRE_MFIP_MFSM_review_date_const, case_count) = trim(MFIP_MFSM_review_date)
                                                
                                                'Navigate to STAT/REVW to confirm review date there
                                                EMWriteScreen "STAT", 20, 13
                                                Call write_value_and_transmit("REVW", 20, 71)
                                                
                                                EmReadScreen REVW_panel_check, 4, 2, 46
                                                ' If REVW_panel_check <> "REVW" Then msgbox "Testing -- 4573 Error unable to reach STAT/REVW"
                                                
                                                'Open the CASH/GRH window
                                                Call write_value_and_transmit("X", 5, 35)
                                                'Read eligibility review date 
                                                EMReadScreen MFIP_STAT_REVW_review_date, 8, 9, 64
                                                'If the review date is blank, then the case should be flagged and skipped for processing
                                                If Instr(MFIP_STAT_REVW_review_date, "_") Then
                                                    If activate_msg_boxes = True Then msgbox "Testing -- error, review date on STAT/REVW for MFIP is empty 4581"
                                                    HIRE_case_details_array(HIRE_MFIP_MFSM_review_date_const, case_count) = trim(MFIP_MFSM_review_date)
                                                    If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                        HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; MFIP - ER Report Date is blank on STAT/REVW"
                                                    Else
                                                        HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "MFIP - ER Report Date is blank on STAT/REVW"
                                                    End If
                                                    HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                Else
                                                    'ER Report date is filled out
                                                    'Convert to MM/DD/YY
                                                    MFIP_STAT_REVW_review_date = replace(MFIP_STAT_REVW_review_date, " ", "/")

                                                    'Update the array
                                                    HIRE_case_details_array(HIRE_MFIP_STAT_REVW_review_date_const, case_count) = trim(MFIP_STAT_REVW_review_date)

                                                    'Compare the review date from MFSM and from STAT/REVW to identify any discrepancies
                                                    MFIP_MFSM_review_date = dateadd("d", 0, MFIP_MFSM_review_date)      'Convert to date
                                                    MFIP_STAT_REVW_review_date = dateadd("d", 0, MFIP_STAT_REVW_review_date)      'Convert to date
                                                    If MFIP_STAT_REVW_review_date <> MFIP_MFSM_review_date Then
                                                        If activate_msg_boxes = True Then msgbox "Testing -- STAT/REVW does not match MFSM 4601"
                                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; Eligibility Review Date on MFSM does not match ER Report Date on STAT/REVW"
                                                        Else
                                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "Eligibility Review Date on MFSM does not match ER Report Date on STAT/REVW"
                                                        End If
                                                        HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                    End If
                                                    If MFIP_STAT_REVW_review_date = MFIP_MFSM_review_date Then
                                                        If DateDiff("m", footer_month_day_year, MFIP_STAT_REVW_review_date) < 0 Then
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review date is before CM and has not been updated correctly 4628"
                                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; MFIP Review Date is prior to current month. Case should be reviewed."
                                                            Else
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "MFIP Review Date is prior to current month. Case should be reviewed."
                                                            End If
                                                            HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                        End If
                                                    End If
                                                End If

                                                'Close the CASH screen
                                                transmit
                                            End If
                                        End If
                                    End If

                                    If GA_active = True Then
                                        'To ensure starting from DAIL, PF3 to get back to DAIL then navigate back to CASE/CURR
                                        'Back to DAIL
                                        PF3
                                        'Navigate back to CASE/CURR
                                        Call write_value_and_transmit("H", dail_row, 3)
                                        'Update the footer month/year and then navigate to ELIG/GA

                                        'Ensure that we are viewing ELIG/FS for the current month, not the dail message month
                                        EMWriteScreen MAXIS_footer_month, 20, 54
                                        EMWriteScreen MAXIS_footer_year, 20, 57
                                        
                                        'Navigate to ELIG/GA from CASE/CURR
                                        EMWriteScreen "ELIG", 20, 22
                                        Call write_value_and_transmit("GA  ", 20, 69)

                                        EMReadScreen no_GA, 10, 24, 2
                                        If no_GA = "NO VERSION" then						'NO GA version means no determination
                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; No version of GA exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            Else
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "No version of GA exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                            End If
                                            HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                        Else

                                            EMWriteScreen "99", 20, 78
                                            transmit
                                            'This brings up the GA versions of eligibility results to search for approved versions
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
                                                If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; No approved eligibility results for GA exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                Else
                                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "No approved eligibility results for GA exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                                End If
                                                HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                            Elseif app_status = "APPROVED" then
                                                'Check for earned income
                                                EMReadScreen vers_number, 1, status_row, 23
                                                Call write_value_and_transmit(vers_number, 18, 54)
                                                'Navigate to GAB1 panel to verify earned income total
                                                Call write_value_and_transmit("GAB1", 20, 70)
                                                EmReadScreen GAB1_panel_check, 4, 3, 49
                                                If GAB1_panel_check <> "GAB1" Then msgbox "Testing -- 4668 Error unable to reach GAB1"
                                                EMReadScreen GA_earned_income_total, 9, 9, 30
                                                GA_earned_income_total = trim(GA_earned_income_total)
                                                'Update array
                                                HIRE_case_details_array(HIRE_GA_earned_income_const, case_count) = GA_earned_income_total

                                                'Navigate to GASM panel
                                                Call write_value_and_transmit("GASM", 20, 71)
                                                
                                                'Read GASM panel to ensure that HRF Reporting is NON-HRF and Budget Cycle is PROSP
                                                EMReadScreen GA_reporting_status, 7, 8, 32
                                                HIRE_case_details_array(HIRE_GA_reporting_status_const, case_count) = trim(GA_reporting_status)
                                                
                                                EMReadScreen GA_GASM_review_date, 8, 11, 32
                                                HIRE_case_details_array(HIRE_GA_GASM_review_date_const, case_count) = trim(GA_GASM_review_date)
                                                
                                                EMReadScreen GA_budget_cycle, 5, 12, 32
                                                HIRE_case_details_array(HIRE_GA_budget_cycle_const, case_count) = trim(GA_budget_cycle)

                                                'Navigate to STAT/REVW to confirm review date there
                                                EMWriteScreen "STAT", 20, 20
                                                Call write_value_and_transmit("REVW", 20, 70)
                                                
                                                EmReadScreen REVW_panel_check, 4, 2, 46
                                                ' If REVW_panel_check <> "REVW" Then msgbox "Testing -- 4692 Error unable to reach STAT/REVW"

                                                'Open the CASH/GRH window
                                                Call write_value_and_transmit("X", 5, 35)
                                                'Read eligibility review date 
                                                EMReadScreen GA_STAT_REVW_review_date, 8, 9, 64
                                                'If the review date is blank, then the case should be flagged and skipped for processing
                                                If Instr(GA_STAT_REVW_review_date, "_") Then
                                                    If activate_msg_boxes = True Then msgbox "Testing -- error, review date on STAT/REVW for GA is empty 4700"
                                                    HIRE_case_details_array(HIRE_GA_GASM_review_date_const, case_count) = trim(GA_STAT_REVW_review_date)
                                                    If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                        HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; GA - ER Report Date is blank on STAT/REVW"
                                                    Else
                                                        HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "GA - ER Report Date is blank on STAT/REVW"
                                                    End If
                                                    HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                Else
                                                    'ER Report date is filled out
                                                    'Convert to MM/DD/YY
                                                    GA_STAT_REVW_review_date = replace(GA_STAT_REVW_review_date, " ", "/")

                                                    'Update the array
                                                    HIRE_case_details_array(HIRE_GA_STAT_REVW_review_date_const, case_count) = trim(GA_STAT_REVW_review_date)

                                                    'Compare the review date from MFSM and from STAT/REVW to identify any discrepancies
                                                    GA_GASM_review_date = dateadd("d", 0, GA_GASM_review_date)      'Convert to date
                                                    GA_STAT_REVW_review_date = dateadd("d", 0, GA_STAT_REVW_review_date)      'Convert to date
                                                    If GA_STAT_REVW_review_date <> GA_GASM_review_date Then
                                                        If activate_msg_boxes = True Then msgbox "Testing -- STAT/REVW does not match GASM 4720"
                                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; Eligibility Review Date on GASM does not match ER Report Date on STAT/REVW"
                                                        Else
                                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "Eligibility Review Date on GASM does not match ER Report Date on STAT/REVW"
                                                        End If
                                                        HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                    End If
                                                    If GA_STAT_REVW_review_date = GA_GASM_review_date Then
                                                        If DateDiff("m", GA_STAT_REVW_review_date, footer_month_day_year) > 0 AND DateDiff("m", GA_GASM_review_date, footer_month_day_year) > 0 Then
                                                            If activate_msg_boxes = True Then msgbox "Testing -- Review dates are before CM 4774"
                                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; GA Review Date is prior to current month. Case should be reviewed"
                                                            Else
                                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "GA Review Date is prior to current month. Case should be reviewed"
                                                            End If
                                                            HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                                        End If
                                                    End If
                                                End If

                                                'Close the CASH screen
                                                transmit
                                            End If
                                        End If
                                    End If

                                Else
                                    'Case is not processable. Write information to array accordingly
                                    HIRE_case_details_array(HIRE_reporting_status_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_sr_report_date_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_recertification_date_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_MFIP_MFSM_review_date_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_MFIP_STAT_REVW_review_date_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_GA_reporting_status_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_GA_budget_cycle_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_GA_earned_income_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_GA_GASM_review_date_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_GA_STAT_REVW_review_date_const, case_count) = "N/A"
                                    HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = "Not processable"
                                    HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                                End If
                            End If    

                            'Only need to check if case is processable if it has not already been determined to be not processable
                            If HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) <> False or trim(HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count)) = "" Then
                                'Handling for SNAP, check if SNAP is active, if it is then verify it meets criteria
                                If SNAP_active = True Then
                                    If HIRE_case_details_array(HIRE_snap_type_const, case_count) = "SNAP" Then
                                        If HIRE_case_details_array(HIRE_snap_status_const, case_count) <> "ACTIVE" OR HIRE_case_details_array(HIRE_reporting_status_const, case_count) <> "SIX MONTH" OR renewal_6_month_check <> True then
                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; SNAP Not Processable"
                                            Else
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "SNAP Not Processable"
                                            End If
                                        End If            
                                    ElseIf HIRE_case_details_array(HIRE_snap_type_const, case_count) = "UHFS" Then
                                        If HIRE_case_details_array(HIRE_snap_status_const, case_count) <> "ACTIVE" OR HIRE_case_details_array(HIRE_reporting_status_const, case_count) <> "SIX MONTH" then
                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; UHFS Not Processable"
                                            Else
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "UHFS Not Processable"
                                            End If
                                        End If
                                    Else
                                        msgbox "Testing -- 4772 missing some handling here. Shouldn't be hitting these, right?"
                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; SNAP or UHFS Not Processable"
                                        Else
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "SNAP or UHFS Not Processable"
                                        End If
                                    End If
                                End If

                                If MFIP_active = True Then
                                    If HIRE_case_details_array(HIRE_MFIP_status_const, case_count) <> "ACTIVE" then
                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; MFIP Not Processable"
                                        Else
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "MFIP Not Processable"
                                        End If
                                    End If
                                End If

                                If GA_active = True Then
                                    If HIRE_case_details_array(HIRE_GA_status_const, case_count) <> "ACTIVE" then
                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; GA Not Processable"
                                        Else
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "GA Not Processable"
                                        End If
                                    ElseIf HIRE_case_details_array(HIRE_GA_status_const, case_count) = "ACTIVE" then
                                        If HIRE_case_details_array(HIRE_GA_reporting_status_const, case_count) <> "NON-HRF" OR HIRE_case_details_array(HIRE_GA_budget_cycle_const, case_count) <> "PROSP" OR HIRE_case_details_array(HIRE_GA_earned_income_const, case_count) < 100 Then 
                                            If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; GA Not Processable"
                                            Else
                                                HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "GA Not Processable"
                                            End If
                                        End If
                                    Else
                                        msgbox "Testing -- 4807 missing some handling here. Shouldn't be hitting these, right?"
                                        If HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) <> "" Then 
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "; ERROR! GA Not Processable"
                                        Else
                                            HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count) & "ERROR! GA Not Processable"
                                        End If
                                    End If
                                End If
                            End If

                            If trim(HIRE_case_details_array(HIRE_case_processing_notes_const, case_count)) <> "" Then
                                HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = False
                            ElseIf trim(HIRE_case_details_array(HIRE_case_processing_notes_const, case_count)) = "" Then
                                HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = True
                            End If

                            'Activate the case details sheet
                            objExcel.Worksheets("Case Details").Activate

                            'Update the Case Details sheet with case data
                            objExcel.Cells(case_excel_row, 1).Value = HIRE_case_details_array(HIRE_case_maxis_case_number_const, case_count)
                            objExcel.Cells(case_excel_row, 2).Value = HIRE_case_details_array(HIRE_case_worker_const, case_count)
                            objExcel.Cells(case_excel_row, 3).Value = HIRE_case_details_array(HIRE_active_programs_const, case_count)
                            objExcel.Cells(case_excel_row, 4).Value = HIRE_case_details_array(HIRE_pending_programs_const, case_count)
                            objExcel.Cells(case_excel_row, 5).Value = HIRE_case_details_array(HIRE_snap_status_const, case_count)
                            objExcel.Cells(case_excel_row, 6).Value = HIRE_case_details_array(HIRE_snap_type_const, case_count)
                            objExcel.Cells(case_excel_row, 7).Value = HIRE_case_details_array(HIRE_reporting_status_const, case_count)
                            objExcel.Cells(case_excel_row, 8).Value = HIRE_case_details_array(HIRE_sr_report_date_const, case_count)
                            objExcel.Cells(case_excel_row, 9).Value = HIRE_case_details_array(HIRE_recertification_date_const, case_count)
                            objExcel.Cells(case_excel_row, 10).Value = HIRE_case_details_array(HIRE_MFIP_status_const, case_count)
                            objExcel.Cells(case_excel_row, 11).Value = HIRE_case_details_array(HIRE_MFIP_MFSM_review_date_const, case_count)
                            objExcel.Cells(case_excel_row, 12).Value = HIRE_case_details_array(HIRE_MFIP_STAT_REVW_review_date_const, case_count)
                            objExcel.Cells(case_excel_row, 13).Value = HIRE_case_details_array(HIRE_GA_status_const, case_count)
                            objExcel.Cells(case_excel_row, 14).Value = HIRE_case_details_array(HIRE_GA_reporting_status_const, case_count)
                            objExcel.Cells(case_excel_row, 15).Value = HIRE_case_details_array(HIRE_GA_budget_cycle_const, case_count)
                            objExcel.Cells(case_excel_row, 16).Value = HIRE_case_details_array(HIRE_GA_earned_income_const, case_count)
                            objExcel.Cells(case_excel_row, 17).Value = HIRE_case_details_array(HIRE_GA_GASM_review_date_const, case_count)
                            objExcel.Cells(case_excel_row, 18).Value = HIRE_case_details_array(HIRE_GA_STAT_REVW_review_date_const, case_count)
                            objExcel.Cells(case_excel_row, 19).Value = HIRE_case_details_array(HIRE_case_processing_notes_const, case_count)
                            objExcel.Cells(case_excel_row, 20).Value = HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count)

                            If HIRE_case_details_array(HIRE_processable_based_on_case_const, case_count) = True and activate_msg_boxes = True Then msgbox "Delete after testing -- Script found case that is in-scope, double-check spreadsheet"
                            
                            'Increment to get to next excel row
                            case_excel_row = case_excel_row + 1

                            EmReadScreen case_curr_check, 4, 2, 55
                            If case_curr_check = "CURR" Then
                                EMWriteScreen MAXIS_footer_month, 20, 54
                                EMWriteScreen MAXIS_footer_year, 20, 57
                                'PF3 back to DAIL
                                PF3 
                            Else
                                'Return to DAIL by PF3
                                PF3

                                'Reset the footer month/year to CM through CASE/CURR
                                Call write_value_and_transmit("H", dail_row, 3)
                                EMWriteScreen MAXIS_footer_month, 20, 54
                                EMWriteScreen MAXIS_footer_year, 20, 57
                                PF3
                            End If
                        
                            'Increment the case_count for updating the array
                            case_count = case_count + 1
                            'Subtract one from dail_row so that the dail_row restarts evaluation of cases now with case details
                            dail_row = dail_row - 1
                        
                        Else
                            'If the MAXIS case number IS in the list of all case numbers, then it is not a new case number and no case details need to be gathered. It can work off the already collected case details.

                            'Before determining whether the DAIL is processable, script determines if it has encountered this DAIL message previously. Based on determination, it then processes (deletes) the dail, skips it, or makes processable determination

                            'Resetting the full_dail_msg to ensure it is not carrying forward to subsequent loops
                            full_dail_msg = ""
                            full_dail_date_hired = ""
                            full_dail_state = ""

                            'Script opens the entire DAIL message to evaluate if it is a new message or not
                            Call write_value_and_transmit("X", dail_row, 3)

                            'Delete after testing - trying to figure out when and why script sometimes does not clear the X
                            EmReadScreen multiple_selections_error_check, 20, 24, 2
                            If InStr(multiple_selections_error_check, "YOU MAY ONLY SELECT") Then msgbox "5006 It failed to clear the previous X"

                            'Handling for reading full dail message depends on message type

                            If dail_type = "HIRE" Then
                                ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                EMReadScreen full_dail_msg_case_number, 35, 6, 44
                                full_dail_msg_case_number = trim(full_dail_msg_case_number)
                                EMReadScreen full_dail_msg_case_number_only, 12, 6, 57
                                full_dail_msg_case_number_only = trim(full_dail_msg_case_number_only)
                                EMReadScreen full_dail_msg_case_name, 35, 7, 44
                                full_dail_msg_case_name = trim(full_dail_msg_case_name)

                                EMReadScreen full_dail_msg_line_1, 60, 9, 5
                                full_dail_msg_line_1 = trim(full_dail_msg_line_1)
                                EMReadScreen full_dail_msg_line_2, 60, 10, 5
                                full_dail_msg_line_2 = trim(full_dail_msg_line_2)
                                EMReadScreen full_dail_msg_line_3, 60, 11, 5
                                full_dail_msg_line_3 = trim(full_dail_msg_line_3)
                                EMReadScreen full_dail_msg_line_4, 60, 12, 5
                                full_dail_msg_line_4 = trim(full_dail_msg_line_4)

                                full_dail_msg = trim(full_dail_msg_case_number & " " & full_dail_msg_case_name & " " & full_dail_msg_line_1 & " " & full_dail_msg_line_2 & " " & full_dail_msg_line_3 & " " & full_dail_msg_line_4)

                                'Read NDNH message employer
                                row = 1
                                col = 1
                                EMSearch "EMPLOYER: ", row, col
                                EMReadScreen full_dail_employer_full_name, 20, row, col + 10
                                full_dail_employer_full_name = trim(full_dail_employer_full_name)

                                If InStr(full_dail_msg_line_1, "NDNH") Then
                                    'Read the NDNH message to find the date hired and convert to MM/DD/YY format
                                    row = 1
                                    col = 1
                                    EMSearch "DATE HIRED   :", row, col
                                    EMReadScreen full_dail_date_hired, 10, row, col + 15
                                    full_dail_date_hired = trim(full_dail_date_hired)
                                    If len(full_dail_date_hired) <> 10 then MsgBox "it is not a 10 character date format"
                                    full_dail_date_hired = Left(full_dail_date_hired, 6) & Right(full_dail_date_hired, 2)

                                    'Read the state of employment
                                    row = 1
                                    col = 1
                                    EMSearch "NDNH MEMB", row, col
                                    EMReadScreen full_dail_state, 2, row, col + 17
                                    full_dail_state = trim(full_dail_state)

                                Else
                                    
                                End If

                                'Transmit back to dail
                                transmit

                            Else
                                MsgBox "Testing -- Dail type is not HIRE. Something went wrong. Dail type is " & dail_type
                            End If

                            'Confirming that dail message lists are updating properly

                            'The script has the full DAIL message and can compare against delete and skip lists to determine if it is a new message

                            If Instr(list_of_DAIL_messages_to_delete_NDNH_known, "*" & full_dail_msg & "*") Then
                                'If the full dail message is within the list of dail messages to delete then the message should be deleted

                                'Delete after testing new functionality
                                ' msgbox "Delete after testing -- Message is within list_of_DAIL_messages_to_delete_NDNH_known"
                                If activate_msg_boxes = True then MsgBox "Testing -- Messages is within list_of_DAIL_messages_to_delete_NDNH_known"

                                ' Resetting variables so they do not carry forward
                                last_dail_check = ""
                                other_worker_error = ""
                                total_dail_msg_count_before = ""
                                total_dail_msg_count_after = ""
                                all_done = ""
                                final_dail_error = ""
                                hire_match = ""
                                
                                'Navigate to INFC
                                If activate_msg_boxes = True then msgbox "testing -- Navigate to INFC"
                                Call write_value_and_transmit("I", dail_row, 3)
                                EMReadScreen SSN_present_check, 9, 3, 63
                                If SSN_present_check = "_________" Then script_end_procedure("Testing -- The script will end because there is a missing SSN. This means handling is needed for these situations.")

                                'Navigate to HIRE interface
                                Call write_value_and_transmit("HIRE", 20, 71)
                                If activate_msg_boxes = True then MsgBox "Testing -- Navigate to HIRE INFC"

                                EMReadScreen infc_hire_check, 8, 2, 50
                                If InStr(infc_hire_check, "HIRE") = 0 Then MsgBox "Testing -- Stop here. Not at INFC/HIRE"

                                'checking for IRS non-disclosure agreement.
                                EMReadScreen agreement_check, 9, 2, 24
                                IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

                                'Navigate through the interface panel to find the matching employer
                                row = 9
                                DO
                                    EMReadScreen infc_case_number, 8, row, 5
                                    infc_case_number = trim(infc_case_number)
                                    IF infc_case_number = full_dail_msg_case_number_only THEN
                                        EMReadScreen infc_employer, 20, row, 36
                                        infc_employer = trim(infc_employer)
                                        IF infc_employer = full_dail_employer_full_name THEN
                                            EMReadScreen known_by_agency, 1, row, 61
                                            IF known_by_agency = " " THEN
                                                EmReadscreen infc_hire_date, 8, row, 20
                                                EmReadscreen infc_hire_state, 2, row, 31
                                                infc_hire_state = trim(infc_hire_state)
                                                If infc_hire_state = "" Then
                                                    If infc_hire_date = full_dail_date_hired Then
                                                        hire_match = TRUE
                                                        match_row = row
                                                        EXIT DO
                                                    End IF
                                                ElseIf infc_hire_state <> "" Then
                                                    If infc_hire_state = full_dail_state AND infc_hire_date = full_dail_date_hired Then
                                                        hire_match = TRUE
                                                        match_row = row
                                                        EXIT DO
                                                    End If
                                                End If
                                            END IF
                                        END IF
                                    END IF
                                    row = row + 1
                                    IF row = 19 THEN
                                        PF8
                                        EmReadscreen end_of_list, 9, 24, 14
                                        If end_of_list = "LAST PAGE" Then Exit Do
                                        row = 9
                                    END IF
                                LOOP UNTIL infc_case_number = ""
                                
                                IF hire_match <> TRUE THEN 
                                    MsgBox "Testing -- No match found in INFC/HIRE"
                                    'The total DAILs decreased by 1, message deleted successfully
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "INFC message unsuccessfully cleared. Validate manually and check if CASE/NOTE added and JOBS panels added. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    script_end_procedure_with_error_report("Testing -- INFC message unsuccessfully cleared. Validate manually and check if CASE/NOTE added and JOBS panels added - something went wrong with clearing the INFC.")
                                ElseIf hire_match = TRUE Then
                                    'entering the INFC/HIRE match '
                                    Call write_value_and_transmit("U", match_row, 3)
                                    EMReadscreen panel_check, 4, 2, 49
                                    IF panel_check <> "NHMD" THEN msgbox "Testing -- We did not enter to clear the match. STOP HERE!!!"
                                    EMWriteScreen "Y", 16, 54
                                    'Agency action must be blank
                                    ' EMWriteScreen "NA", 17, 54
                                    If activate_msg_boxes = True then MsgBox "Testing -- Validate that correct information has been written to the case! Script is about to save update INFC update. STOP here if needed."
                                    TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
                                    TRANSMIT 'this confirms the cleared status'
                                    PF3
                                    EMReadscreen cleared_confirmation, 1, match_row, 61
                                    IF cleared_confirmation = " " THEN 
                                        MsgBox "Testing -- the match did not appear to clear"
                                        'The total DAILs decreased by 1, message deleted successfully
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "INFC message unsuccessfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with clearing the INFC message at line 3884.")
                                    ElseIf cleared_confirmation <> " " THEN 
                                        If activate_msg_boxes = True then MsgBox "Testing -- the match appears to have cleared. Verify manually before continuing"
                                        'The total DAILs decreased by 1, message deleted successfully
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "INFC message successfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    End If
                                End If

                                PF3' this takes us back to DAIL/DAIL

                                Call nav_back_to_dail_check(True)

                                EMReadScreen infc_clear_error, 40, 24, 2
                                infc_clear_error = trim(infc_clear_error)
                                If Instr(infc_clear_error, "THIS IS NOT YOUR DAIL REPORT") = 0 Then MsgBox "Testing -- Stop here. Something happened after clearing the INFC 5057"
                                
                                If activate_msg_boxes = True then MsgBox "The message has been deleted. Did anything go wrong? If so, stop here!"
                            ElseIf Instr(list_of_DAIL_messages_to_delete_NDNH_not_known, "*" & full_dail_msg & "*") Then
                                'If the full dail message is within the list of dail messages to delete then the message should be deleted

                                'Delete after testing new functionality
                                ' msgbox "Delete after testing -- Message is within list_of_DAIL_messages_to_delete_NDNH_not_known"
                                If activate_msg_boxes = True then MsgBox "Testing -- Message within list_of_DAIL_messages_to_delete_NDNH_not_known"

                                ' Resetting variables so they do not carry forward
                                last_dail_check = ""
                                other_worker_error = ""
                                total_dail_msg_count_before = ""
                                total_dail_msg_count_after = ""
                                all_done = ""
                                final_dail_error = ""
                                hire_match = ""
                                
                                'Navigate to INFC
                                If activate_msg_boxes = True then msgbox "Testing -- Navigate to INFC"
                                Call write_value_and_transmit("I", dail_row, 3)
                                EMReadScreen SSN_present_check, 9, 3, 63
                                If SSN_present_check = "_________" Then script_end_procedure("Testing -- The script will end because there is a missing SSN. This means handling is needed for these situations.")

                                'Navigate to HIRE interface
                                If activate_msg_boxes = True then msgbox "Testing -- Navigate to HIRE INFC"
                                Call write_value_and_transmit("HIRE", 20, 71)

                                EMReadScreen infc_hire_check, 8, 2, 50
                                If InStr(infc_hire_check, "HIRE") = 0 Then MsgBox "Testing -- Stop here. Not at INFC/HIRE"

                                'checking for IRS non-disclosure agreement.
                                EMReadScreen agreement_check, 9, 2, 24
                                IF agreement_check = "Automated" THEN script_end_procedure("Testing -- To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

                               'Navigate through the interface panel to find the matching employer
                                row = 9
                                DO
                                    EMReadScreen infc_case_number, 8, row, 5
                                    infc_case_number = trim(infc_case_number)
                                    IF infc_case_number = full_dail_msg_case_number_only THEN
                                        EMReadScreen infc_employer, 20, row, 36
                                        infc_employer = trim(infc_employer)
                                        IF infc_employer = full_dail_employer_full_name THEN
                                            EMReadScreen known_by_agency, 1, row, 61
                                            IF known_by_agency = " " THEN
                                                EmReadscreen infc_hire_date, 8, row, 20
                                                EmReadscreen infc_hire_state, 2, row, 31
                                                infc_hire_state = trim(infc_hire_state)
                                                If infc_hire_state = "" Then
                                                    If infc_hire_date = full_dail_date_hired Then
                                                        hire_match = TRUE
                                                        match_row = row
                                                        EXIT DO
                                                    End IF
                                                ElseIf infc_hire_state <> "" Then
                                                    If infc_hire_state = full_dail_state AND infc_hire_date = full_dail_date_hired Then
                                                        hire_match = TRUE
                                                        match_row = row
                                                        EXIT DO
                                                    End If
                                                End If
                                            END IF
                                        END IF
                                    END IF
                                    row = row + 1
                                    IF row = 19 THEN
                                        PF8
                                        EmReadscreen end_of_list, 9, 24, 14
                                        If end_of_list = "LAST PAGE" Then Exit Do
                                        row = 9
                                    END IF
                                LOOP UNTIL infc_case_number = ""
                                
                                IF hire_match <> TRUE THEN 
                                    MsgBox "Testing -- No match found in INFC/HIRE"
                                    'The total DAILs decreased by 1, message deleted successfully
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "INFC message unsuccessfully cleared. Validate manually and check if CASE/NOTE added and JOBS panels added. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    script_end_procedure_with_error_report("Testing -- INFC message unsuccessfully cleared. Validate manually and check if CASE/NOTE added and JOBS panels added - something went wrong with clearing the INFC.")
                                ElseIf hire_match = TRUE Then
                                    'entering the INFC/HIRE match '
                                    Call write_value_and_transmit("U", match_row, 3)
                                    EMReadscreen panel_check, 4, 2, 49
                                    IF panel_check <> "NHMD" THEN msgbox "Testing -- We did not enter to clear the match. STOP HERE!!!"
                                    EMWriteScreen "N", 16, 54
                                    EMWriteScreen "NA", 17, 54
                                    If activate_msg_boxes = True then MsgBox "Testing -- Validate that correct information has been written to the case! Script is about to save update INFC update. STOP here if needed."
                                    TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
                                    TRANSMIT 'this confirms the cleared status'
                                    PF3
                                    EMReadscreen cleared_confirmation, 1, match_row, 61
                                    IF cleared_confirmation = " " THEN 
                                        MsgBox "Testing -- the match did not appear to clear"
                                        'The total DAILs decreased by 1, message deleted successfully
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "INFC message unsuccessfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with clearing the INFC message at line 3884.")
                                    ElseIf cleared_confirmation <> " " THEN 
                                        If activate_msg_boxes = True then msgbox "Testing -- the match appears to have cleared. Verify manually before continuing"
                                        'The total DAILs decreased by 1, message deleted successfully
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "INFC message successfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    End If
                                End If

                                PF3' this takes us back to DAIL/DAIL

                                Call nav_back_to_dail_check(True)

                                EMReadScreen infc_clear_error, 40, 24, 2
                                infc_clear_error = trim(infc_clear_error)
                                If Instr(infc_clear_error, "THIS IS NOT YOUR DAIL REPORT") = 0 Then MsgBox "Testing -- Stop here. Something happened after clearing the INFC 4018"

                                If activate_msg_boxes = True then MsgBox "The message has been deleted. Did anything go wrong? If so, stop here!"
                            ElseIf Instr(list_of_DAIL_messages_to_delete_SDNH, "*" & full_dail_msg & "*") Then

                                'Delete after testing new functionality
                                ' msgbox "Testing -- Script is about to delete the duplicate SDNH message or the reviewed SDNH. Make sure that it is correct!!"
                                'If the full dail message is within the list of dail messages to delete then the message should be deleted
                                If activate_msg_boxes = True then MsgBox "Testing -- Script is about to delete the duplicate SDNH message or the reviewed SDNH. Make sure that it is correct!!"

                                'Resetting variables so they do not carry forward
                                last_dail_check = ""
                                other_worker_error = ""
                                total_dail_msg_count_before = ""
                                total_dail_msg_count_after = ""
                                all_done = ""
                                final_dail_error = ""
                                
                                'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
                                EMReadScreen last_dail_check, 12, 3, 67
                                last_dail_check = trim(last_dail_check)

                                'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
                                last_dail_check = split(last_dail_check, " ")

                                If last_dail_check(0) = last_dail_check(2) then 
                                    'The script is about to delete the LAST message in the DAIL so script will exit do loop after deletion, also works if it is about to delete the ONLY message in the DAIL
                                    all_done = true
                                End If

                                If activate_msg_boxes = True then msgbox "Testing -- Script will now delete the SDNH message"
                                'Delete the message
                                Call write_value_and_transmit("D", dail_row, 3)
                                activate_msg_boxes = False

                                'Handling for deleting message under someone else's x number
                                EMReadScreen other_worker_error, 25, 24, 2
                                other_worker_error = trim(other_worker_error)

                                If other_worker_error = "ALL MESSAGES WERE DELETED" Then
                                    'Script deleted the final message in the DAIL
                                    dail_row = dail_row - 1
                                    dail_msg_deleted_count = dail_msg_deleted_count + 1
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    'Exit do loop as all messages are deleted
                                    all_done = true

                                ElseIf other_worker_error = "" Then
                                    'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion

                                    'Handling to check if message actually deleted
                                    total_dail_msg_count_before = last_dail_check(2) * 1
                                    EMReadScreen total_dail_msg_count_after, 12, 3, 67

                                    total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                    total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                    If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                        'The total DAILs decreased by 1, message deleted successfully
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    Else
                                        'The total DAILs did not decrease by 1, something went wrong
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 881.")
                                    End If

                                ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
                                    
                                    'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
                                    transmit

                                    'Reads the total number of DAILS after deleting to determine if it decreased by 1
                                    EMReadScreen total_dail_msg_count_after, 12, 3, 67

                                    'Checks if final DAIL message deleted
                                    EMReadScreen final_dail_error, 25, 24, 2

                                    If final_dail_error = "ALL MESSAGES WERE DELETED" Then
                                        'All DAIL messages deleted so indicates deletion a success
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        'No more DAIL messages so exit do loop
                                        all_done = True
                                    ElseIf trim(final_dail_error) = "" Then
                                        'Handling to check if message actually deleted
                                        total_dail_msg_count_before = last_dail_check(2) * 1

                                        total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                        total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                        If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                            'The total DAILs decreased by 1, message deleted successfully
                                            dail_row = dail_row - 1
                                            dail_msg_deleted_count = dail_msg_deleted_count + 1
                                            objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully cleared. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        Else
                                            objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                            script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 4030.")
                                        End If

                                    Else
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 4035.")
                                    End if
                                    
                                Else
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 4040.")
                                End If

                                If activate_msg_boxes = True then MsgBox "The message has been deleted. Did anything go wrong? If so, stop here!"
                            ElseIf Instr(list_of_DAIL_messages_to_skip, "*" & full_dail_msg & "*") Then
                                'If the full message is on the list of dail messages to skip then the message should be skipped

                                If Instr(DAIL_messages_to_skip, "*" & full_dail_msg & "*") and activate_msg_boxes = True then msgbox "It hit a message and skipped it that was on list of skips > " & full_dail_msg

                            ElseIf Instr(list_of_DAIL_messages_to_delete_NDNH_known, "*" & full_dail_msg & "*") = 0 AND Instr(list_of_DAIL_messages_to_delete_NDNH_not_known, "*" & full_dail_msg & "*") = 0 AND Instr(list_of_DAIL_messages_to_delete_SDNH, "*" & full_dail_msg & "*") = 0 AND Instr(list_of_DAIL_messages_to_skip, "*" & full_dail_msg & "*") = 0 Then
                                'If the full dail message is NOT in the list of dail messages to delete AND the full dail messages is NOT in the list of skip messages then it SHOULD be a new dail message and therefore it needs to be evaluated

                                'Gather details on DAIL message, should capture DAIL details in spreadsheet even if ultimately not actionable
                            
                                'Reset the array
                                ReDim Preserve DAIL_message_array(DAIL_excel_row_const, dail_count)
                                DAIL_message_array(dail_maxis_case_number_const, DAIL_count) = MAXIS_case_number
                                DAIL_message_array(dail_worker_const, DAIL_count) = worker

                                'Use for next loop to match the individual DAIL message to the corresponding array item of matching Case Details
                                for each_case = 0 to UBound(HIRE_case_details_array, 2)
                                    'Iterate through each of the cases 
                                    If DAIL_message_array(dail_maxis_case_number_const, dail_count) = HIRE_case_details_array(HIRE_case_maxis_case_number_const, each_case) Then
                                        'As the for to loop iterates through each case details array, if the dail maxis case number for the dail message array matches the maxis case number for the case details array then it can pull the case details from the array  
                                        
                                        'Clearing out process_dail_message
                                        process_dail_message = ""

                                        'Read dail message details
                                        EMReadScreen dail_type, 4, dail_row, 6
                                        dail_type = trim(dail_type)

                                        EMReadScreen dail_month, 8, dail_row, 11
                                        dail_month = trim(dail_month)

                                        EMReadScreen dail_msg, 61, dail_row, 20
                                        dail_msg = trim(dail_msg)

                                        'Update the DAIL message array with details
                                        DAIL_message_array(dail_type_const, dail_count) = dail_type
                                        DAIL_message_array(dail_month_const, dail_count) = dail_month
                                        DAIL_message_array(dail_msg_const, dail_count) = dail_msg
                                        DAIL_message_array(full_dail_msg_const, dail_count) = full_dail_msg

                                        'Activate the DAIL Messages sheet
                                        objExcel.Worksheets("DAIL Messages").Activate

                                        'Write dail details to the Excel sheet
                                        objExcel.Cells(dail_excel_row, 1).Value = DAIL_message_array(dail_maxis_case_number_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 2).Value = DAIL_message_array(dail_worker_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 3).Value = DAIL_message_array(dail_type_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 4).Value = DAIL_message_array(dail_month_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 5).Value = DAIL_message_array(dail_msg_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 6).Value = DAIL_message_array(full_dail_msg_const, dail_count)

                                        If HIRE_case_details_array(HIRE_processable_based_on_case_const, each_case) = False Then
                                            If Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "SR Report Date and Recertification are not 6 months apart") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "SR Report Date and/or Recertification Date is missing") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "SNAP Review Dates are prior to current month. Case should be reviewed") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "MFIP - ER Report Date is blank on STAT/REVW") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "Eligibility Review Date on MFSM does not match ER Report Date on STAT/REVW") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "MFIP Review Date is prior to current month. Case should be reviewed") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "GA - ER Report Date is blank on STAT/REVW") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "GA Review Date is prior to current month. Case should be reviewed") OR _
                                                Instr(HIRE_case_details_array(HIRE_case_processing_notes_const, each_case), "Eligibility Review Date on GASM does not match ER Report Date on STAT/REVW") Then
                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "QI review needed." & HIRE_case_details_array(HIRE_case_processing_notes_const, each_case)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                            Else
                                                DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable based on Case Details: " & HIRE_case_details_array(HIRE_case_processing_notes_const, each_case)
                                                not_processable_msg_count = not_processable_msg_count + 1
                                            End If

                                            'The dail message should not be processed due to case details
                                            process_dail_message = False

                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                            'Activate the DAIL Messages sheet
                                            objExcel.Worksheets("DAIL Messages").Activate

                                            'To do - Delete after testing 
                                            If MAXIS_case_number = "940366" Then msgbox "Delete after testing - Line 5465."

                                            'Update the Excel sheet
                                            objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                        
                                        ElseIf HIRE_case_details_array(HIRE_processable_based_on_case_const, each_case) = True Then     
                                            
                                            'Convert dail month to month day year in a date format
                                            dail_month_day_year = replace(dail_month, " ", "/01/")
                                            dail_month_day_year = dateadd("m", 0, dail_month_day_year)

                                            'Determine if dail month is more than 6 months old
                                            dail_over_6_months_old = datediff("m", dail_month_day_year, footer_month_day_year)

                                            If dail_over_6_months_old > 6 Then
                                                If dail_type = "HIRE" Then
                                                    If activate_msg_boxes = True then msgbox "Testing -- dail is over 6 months old"
                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "Not processable as the DAIL month is over 6 months old. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                    not_processable_msg_count = not_processable_msg_count + 1

                                                    'The dail message cannot be processed as it is over 6 months old
                                                    process_dail_message = False

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                Else
                                                    MsgBox "Testing -- something went wrong around 5374. Wasn't a HIRE type Dail when determining if older than 6 months"
                                                End If

                                                'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                            Else
                                                If HIRE_case_details_array(HIRE_snap_type_const, each_case) = "SNAP" OR HIRE_case_details_array(HIRE_snap_type_const, each_case) = "UHFS" Then
                                                    If DateAdd("m", 0, HIRE_case_details_array(HIRE_recertification_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, HIRE_case_details_array(HIRE_sr_report_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) Then
                                                        If activate_msg_boxes = True then Msgbox "The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"

                                                        If dail_type = "HIRE" Then
                                                            If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 0, footer_month_day_year) Then
                                                                
                                                                DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                                objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                                not_processable_msg_count = not_processable_msg_count + 1

                                                                'The dail message cannot be processed due to timing of recertification or SR report date
                                                                process_dail_message = False

                                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                            Else
                                                                'Process the HIRE message
                                                                process_dail_message = True
                                                            End If
                                                        Else
                                                            'Add handling here if needed
                                                        End If

                                                    Else
                                                        'Situation where it is less than 6 months old AND recert/renewal is not next month
                                                        'If neither the recertification or SR report date is next month then we assume the dail message can be processed since processable based on case details is True. So set the process_dail_message to True to gather more information about the dail message
                                                        process_dail_message = True
                                                    End If
                                                End If

                                                'Handling for MFIP to check if review or recert is CM + 1. If so, checks if DAIL month is CM + 1 too. If that's the case, it will skip processing the message.
                                                If HIRE_case_details_array(HIRE_MFIP_status_const, each_case) = "ACTIVE" Then
                                                    'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                                    'Subtract 6 months from ER Report Date to get review date
                                                    ER_report_minus_6_months = DateAdd("m", -6, HIRE_case_details_array(HIRE_MFIP_STAT_REVW_review_date_const, each_case))

                                                    If DateAdd("m", 0, HIRE_case_details_array(HIRE_MFIP_STAT_REVW_review_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, ER_report_minus_6_months) = DateAdd("m", 1, footer_month_day_year) Then
                                                        ' If activate_msg_boxes = True Then Msgbox "The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"
                                                        ' Msgbox "5537 Delete after testing -- The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"

                                                        If dail_type = "HIRE" Then
                                                            
                                                            If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 0, footer_month_day_year) Then
                                                                ' msgbox "5542 Delete after testing -- Unable to process the message since recert is next month and DAIL month is current month"

                                                                If trim(DAIL_message_array(dail_processing_notes_const, dail_count)) = "" then
                                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal for MFIP. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                                Else
                                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = DAIL_message_array(dail_processing_notes_const, dail_count) & "; Not Processable due to DAIL Month & Recert/Renewal for MFIP. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                                End If
                                                                
                                                                objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                                not_processable_msg_count = not_processable_msg_count + 1

                                                                'The dail message cannot be processed due to timing of recertification or SR report date
                                                                process_dail_message = False
                                                                'Add to skip list
                                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                            Else
                                                                'DAIL message can be processed
                                                                process_dail_message = True
                                                            End If
                                                        End If

                                                    Else
                                                        'If neither the recertification or SR report date is next month then we assume the dail message can be processed since processable based on case details is True. So set the process_dail_message to True to gather more information about the dail message
                                                        process_dail_message = True
                                                        
                                                    End If
                                                End If

                                                'Handling for GA to check if review or recert is CM + 1. If so, checks if DAIL month is CM + 1 too. If that's the case, it will skip processing the message.
                                                If HIRE_case_details_array(HIRE_GA_status_const, each_case) = "ACTIVE" Then
                                                    'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                                    'Subtract 6 months from ER Report Date to get review date

                                                    ER_report_minus_6_months = DateAdd("m", -6, HIRE_case_details_array(HIRE_GA_GASM_review_date_const, each_case))

                                                    If DateAdd("m", 0, HIRE_case_details_array(HIRE_GA_GASM_review_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, ER_report_minus_6_months) = DateAdd("m", 1, footer_month_day_year) Then
                                                        If activate_msg_boxes = True Then Msgbox "The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"
                                                        Msgbox "5579 Delete after testing -- The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"

                                                        If dail_type = "HIRE" Then
                                                            msgbox "5583 Delete after testing -- Unable to process the message since recert is next month and DAIL month is current month"
                                                            
                                                            If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 0, footer_month_day_year) Then

                                                                If trim(DAIL_message_array(dail_processing_notes_const, dail_count)) = "" then
                                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal for GA. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                                Else
                                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = DAIL_message_array(dail_processing_notes_const, dail_count) & "; Not Processable due to DAIL Month & Recert/Renewal for GA. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                                End If
                                                                
                                                                objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                                not_processable_msg_count = not_processable_msg_count + 1

                                                                'The dail message cannot be processed due to timing of recertification or SR report date
                                                                process_dail_message = False
                                                                'Add to skip list
                                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                            Else
                                                                'DAIL message can be processed
                                                                process_dail_message = True
                                                            End If
                                                        End If

                                                    Else
                                                        'If neither the recertification or SR report date is next month then we assume the dail message can be processed since processable based on case details is True. So set the process_dail_message to True to gather more information about the dail message
                                                        process_dail_message = True
                                                        
                                                    End If
                                                End If
                                            End If

                                            If process_dail_message = True and dail_type = "HIRE" Then

                                                If InStr(dail_msg, "NDNH MEMB") Then

                                                    'Reset variables to ensure they don't carry forward through do loop
                                                    HIRE_memb_number = ""
                                                    no_exact_JOBS_panel_matches = ""
                                                    list_of_employers_on_jobs_panels = "*"
                                                    JOBS_footer_month = ""
                                                    JOBS_footer_year = ""
                                                    HIRE_memb_number = ""
                                                    date_hired = ""
                                                    HIRE_employer_name = ""
                                                    NDNH_MAXIS_name = ""
                                                    NDNH_new_hire_name = ""
                                                    hire_message_member_name = ""
                                                    hire_message_case_number = ""
                                                    HIRE_employer_name_first_word = ""
                                                    HIRE_employer_name_TIKL = ""
                                                    tikl_case_number = ""
                                                    tikl_case_name = ""
                                                    blank_state_check = ""

                                                    'Blanking variables to check for potential SNAP income exclusion
                                                    hh_memb_age = ""
                                                    under_18_check = ""
                                                    hh_memb_rel_to_applicant = ""
                                                    child_of_hh_member = ""
                                                    school_status = ""
                                                    school_status_qualifies = ""
                                                    school_type = ""
                                                    school_type_qualifies = ""
                                                    snap_earned_income_minor_exclusion = ""
                                                    fs_eligibility_eligible = ""
                                                    fs_eligibility_status_check = ""

                                                    'Read the HIRE message member name to navigate back if needed
                                                    EMReadScreen hire_message_member_name, 8, dail_row - 1, 5
                                                    EMReadScreen hire_message_case_number, 8, dail_row - 1, 73
                                                    hire_message_case_number = trim(hire_message_case_number)

                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    'Delete after testing - trying to figure out when and why script sometimes does not clear the X
                                                    EmReadScreen multiple_selections_error_check, 20, 24, 2
                                                    If InStr(multiple_selections_error_check, "YOU MAY ONLY SELECT") Then msgbox "5675 It failed to clear the previous X"

                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    EMReadScreen check_full_dail_msg_case_number, 35, 6, 44
                                                    check_full_dail_msg_case_number = trim(check_full_dail_msg_case_number)
                                                    EMReadScreen check_full_dail_msg_case_name, 35, 7, 44
                                                    check_full_dail_msg_case_name = trim(check_full_dail_msg_case_name)

                                                    EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                    check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                    EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                    check_full_dail_msg_line_2 = trim(check_full_dail_msg_line_2)
                                                    EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                    check_full_dail_msg_line_3 = trim(check_full_dail_msg_line_3)
                                                    EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                    check_full_dail_msg_line_4 = trim(check_full_dail_msg_line_4)

                                                    check_full_dail_msg = trim(check_full_dail_msg_case_number & " " & check_full_dail_msg_case_name & " " & check_full_dail_msg_line_1 & " " & check_full_dail_msg_line_2 & " " & check_full_dail_msg_line_3 & " " & check_full_dail_msg_line_4)

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- Something went wrong around 5556. check_full_dail_msg " & check_full_dail_msg & vbNewLine & vbNewLine & " full_dail_msg " & full_dail_msg
                                                    End if

                                                    'Read the Case Name and Case Number to process TIKLs as needed
                                                    EmReadScreen tikl_case_number, 10, 6, 57
                                                    tikl_case_number = trim(tikl_case_number)
                                                    EmReadScreen tikl_case_name, 25, 7, 55
                                                    tikl_case_name = trim(tikl_case_name)

                                                    'Identify where 'NDNH MEMB:' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "NDNH MEMB", row, col
                                                    EMReadScreen blank_state_check, 2, row, col + 17

                                                    'Identify where 'NDNH MEMB:' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "NDNH MEMB", row, col
                                                    EMReadScreen HIRE_memb_number, 2, row, col + 10
                                                    HIRE_memb_number = trim(HIRE_memb_number)
                                                    If HIRE_memb_number = "00" then msgbox "Testing -- HH MEMB 00 - this is an error message. Need handling for this situation"

                                                    'Identify where 'DATE HIRED   :' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "DATE HIRED   :", row, col
                                                    EMReadScreen date_hired, 10, row, col + 15
                                                    date_hired = trim(date_hired)

                                                    If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then
                                                        msgbox "Testing -- date hired is EM or unknown. How to handle?"
                                                    Else
                                                        Call ONLY_create_MAXIS_friendly_date(date_hired)
                                                        date_split = split(date_hired, "/")
                                                        month_hired = date_split(0)
                                                        day_hired = date_split(1)
                                                        year_hired = date_split(2)
                                                    End if

                                                    'Identify where ' Employer:' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "EMPLOYER: ", row, col
                                                    EMReadScreen HIRE_employer_name, 20, row, col + 10
                                                    HIRE_employer_name = trim(HIRE_employer_name)
                                                    EMReadScreen HIRE_employer_name_TIKL, 25, row, col + 10
                                                    HIRE_employer_name_TIKL = TRIM(HIRE_employer_name_TIKL)

                                                    If blank_state_check <> "??" and HIRE_employer_name <> "" Then

                                                        'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
                                                        HIRE_employer_name_split = split(HIRE_employer_name, " ")

                                                        If len(HIRE_employer_name_split(0)) < 4 and Ubound(HIRE_employer_name_split) > 0 Then
                                                            HIRE_employer_name_first_word = HIRE_employer_name_split(0) & " " & HIRE_employer_name_split(1)
                                                            If activate_msg_boxes = True then MsgBox "First word less than 3 characters long. HIRE_employer_name_first_word is " & HIRE_employer_name_first_word  
                                                        Else
                                                            HIRE_employer_name_first_word = HIRE_employer_name_split(0)   
                                                            If activate_msg_boxes = True then MsgBox "First word longer than 3 characters long. HIRE_employer_name_first_word is " & HIRE_employer_name_first_word
                                                        End If

                                                        If instr(len(HIRE_employer_name_first_word), HIRE_employer_name_first_word, ",") = len(HIRE_employer_name_first_word) then 
                                                            HIRE_employer_name_first_word = Mid(HIRE_employer_name_first_word, 1, len(HIRE_employer_name_first_word) - 1)
                                                            If activate_msg_boxes = True then MsgBox "Last character is a comma. HIRE_employer_name_first_word is now " & HIRE_employer_name_first_word
                                                        End If

                                                        'Identify where 'MAXIS NAME   :' text is so that script can account for slight changes in location in MAXIS
                                                        'Set row and col
                                                        row = 1
                                                        col = 1
                                                        EMSearch "MAXIS NAME   :", row, col
                                                        EMReadScreen NDNH_MAXIS_name, 30, row, col + 15
                                                        NDNH_MAXIS_name = trim(NDNH_MAXIS_name)

                                                        'Identify where 'NEW HIRE NAME:' text is so that script can account for slight changes in location in MAXIS
                                                        'Set row and col
                                                        row = 1
                                                        col = 1
                                                        EMSearch "NEW HIRE NAME:", row, col
                                                        EMReadScreen NDNH_new_hire_name, 30, row, col + 15
                                                        NDNH_new_hire_name = trim(NDNH_new_hire_name)

                                                        'Transmit back to DAIL message
                                                        transmit

                                                        EMWriteScreen "S", dail_row, 3
                                                        EMSendKey "<enter>"
                                                        EMReadScreen background_check, 25, 7, 30
                                                        If InStr(background_check, "A Background transaction") Then
                                                            EMWaitReady 2, 2000
                                                            Do
                                                                background_check = ""
                                                                PF3
                                                                EMWaitReady 2, 2000
                                                                EMWriteScreen "S", dail_row, 3
                                                                EMWaitReady 2, 2000
                                                                EMSendKey "<enter>"
                                                                EMWaitReady 2, 2000
                                                                EMReadScreen background_check, 25, 7, 30
                                                                If InStr(background_check, "A Background transaction") = 0 then Exit Do
                                                            Loop
                                                        End If

                                                        EMReadScreen self_panel_check, 4, 2, 50
                                                        If self_panel_check = "SELF" Then
                                                            EMWaitReady 2, 2000
                                                            EMWaitReady 2, 2000
                                                            EMWriteScreen "DAIL", 16, 43
                                                            EMWriteScreen "DAIL", 21, 70
                                                            transmit

                                                            EMReadScreen back_to_dail_check, 8, 1, 72
                                                            If back_to_dail_check = "FMKDLAM6" Then

                                                                'Navigate to CASE/CURR to force DAIL to reset and then PF3 back to get back to start of the DAIL
                                                                Call write_value_and_transmit("H", dail_row, 3)
                                                                PF3

                                                                'Reset DAIL to only HIRE messages
                                                                Call write_value_and_transmit("X", 4, 12)
                                                                EMWriteScreen "_", 7, 39
                                                                Call write_value_and_transmit("X", 13, 39)

                                                                'Script should now navigate to specific member name, or at least get close
                                                                EMWriteScreen hire_message_member_name, 21, 25
                                                                transmit

                                                                'Script will enter do loop to find match

                                                                Do
                                                                    return_full_dail_msg = ""
                                                                    return_full_dail_msg_case_number = ""
                                                                    return_full_dail_msg_case_name = ""
                                                                    return_full_dail_msg_line_1 = ""
                                                                    return_full_dail_msg_line_2 = ""
                                                                    return_full_dail_msg_line_3 = ""
                                                                    return_full_dail_msg_line_4 = ""

                                                                    'Enters “X” on DAIL message to open full message. 
                                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                                    'Delete after testing - trying to figure out when and why script sometimes does not clear the X
                                                                    EmReadScreen multiple_selections_error_check, 20, 24, 2
                                                                    If InStr(multiple_selections_error_check, "YOU MAY ONLY SELECT") Then msgbox "5844 It failed to clear the previous X"

                                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                                    EMReadScreen return_full_dail_msg_case_number, 35, 6, 44
                                                                    return_full_dail_msg_case_number = trim(return_full_dail_msg_case_number)
                                                                    EMReadScreen return_full_dail_msg_case_name, 35, 7, 44
                                                                    return_full_dail_msg_case_name = trim(return_full_dail_msg_case_name)

                                                                    EMReadScreen return_full_dail_msg_line_1, 60, 9, 5
                                                                    return_full_dail_msg_line_1 = trim(return_full_dail_msg_line_1)
                                                                    EMReadScreen return_full_dail_msg_line_2, 60, 10, 5
                                                                    return_full_dail_msg_line_2 = trim(return_full_dail_msg_line_2)
                                                                    EMReadScreen return_full_dail_msg_line_3, 60, 11, 5
                                                                    return_full_dail_msg_line_3 = trim(return_full_dail_msg_line_3)
                                                                    EMReadScreen return_full_dail_msg_line_4, 60, 12, 5
                                                                    return_full_dail_msg_line_4 = trim(return_full_dail_msg_line_4)

                                                                    return_full_dail_msg = trim(return_full_dail_msg_case_number & " " & return_full_dail_msg_case_name & " " & return_full_dail_msg_line_1 & " " & return_full_dail_msg_line_2 & " " & return_full_dail_msg_line_3 & " " & return_full_dail_msg_line_4)

                                                                    If return_full_dail_msg = check_full_dail_msg Then 
                                                                        transmit
                                                                        Exit Do
                                                                    Else
                                                                        transmit
                                                                        dail_row = dail_row + 1

                                                                        'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                                                                        EMReadScreen new_case, 8, dail_row, 63
                                                                        new_case = trim(new_case)
                                                                        IF new_case <> "CASE NBR" THEN 
                                                                            'If there is NOT a new case number, the script will top the message
                                                                            Call write_value_and_transmit("T", dail_row, 3)
                                                                        ELSEIF new_case = "CASE NBR" THEN
                                                                            'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                                                                            Call write_value_and_transmit("T", dail_row + 1, 3)
                                                                        End if
                                                                    End If

                                                                Loop

                                                                'Reset the dail_row back to 6
                                                                dail_row = 6

                                                                EMWriteScreen "S", dail_row, 3
                                                                EMSendKey "<enter>"
                                                            Else

                                                                'Initial dialog - select whether to create a list or process a list
                                                                Dialog1 = ""
                                                                BeginDialog Dialog1, 0, 0, 306, 220, "Unable to return to DAIL. Double-check the issue."
                                                                
                                                                ButtonGroup ButtonPressed
                                                                    OkButton 205, 200, 40, 15
                                                                    CancelButton 245, 200, 40, 15
                                                                EndDialog

                                                                Do
                                                                    Dialog Dialog1
                                                                Loop until ButtonPressed = OK

                                                            End If
                                                        End If

                                                        EMWriteScreen "MEMB", 20, 71
                                                        Call write_value_and_transmit(HIRE_memb_number, 20, 76)

                                                        EMReadScreen memb_panel_check, 4, 2, 48
                                                        IF memb_panel_check <> "MEMB" Then 
                                                            EMReadScreen summ_panel_check, 4, 2, 46
                                                            If summ_panel_check = "SUMM" Then
                                                                EMWriteScreen "MEMB", 20, 71
                                                                Call write_value_and_transmit(HIRE_memb_number, 20, 76)
                                                                EMReadScreen memb_panel_check, 4, 2, 48
                                                                IF memb_panel_check <> "MEMB" Then MsgBox "Testing -- second attempt to get to MEMB failed 5709"
                                                            Else
                                                                MsgBox "Testing -- not on Summ 4830. Will attempt to go back to DAIL"
                                                            End If
                                                        End If

                                                        'Ensure the script is not creating a new MEMB panel
                                                        EMReadScreen new_memb_panel_check, 12, 8, 22
                                                        If new_memb_panel_check = "Arrival Date" Then
                                                            PF3
                                                            PF10
                                                            msgbox "Testing -- Script tried to navigate to a HH Memb that doesn't exist. It should have deleted the panel but double check MAKE SURE IT DELETED ADDED PANEL"

                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "HIRE message identifies a HH Member that does not exist on the case (" & HIRE_memb_number & "). Review needed." & " Message should not be deleted."
                                                        Else
                                                        
                                                            'Check the HH Memb's age and relationship status
                                                            EMReadScreen hh_memb_age, 2, 8, 76
                                                            hh_memb_age = trim(hh_memb_age)
                                                            'Convert age into a number
                                                            If hh_memb_age = "" then MsgBox "Testing -- No age on panel. stop here"
                                                            If hh_memb_age <> "" Then hh_memb_age = hh_memb_age * 1

                                                            If hh_memb_age > 17 then 
                                                                under_18_check = False 
                                                            Else
                                                                under_18_check = True
                                                            End If

                                                            'Convert age to a number
                                                            EMReadScreen hh_memb_rel_to_applicant, 2, 10, 42
                                                            If hh_memb_rel_to_applicant = "03" OR hh_memb_rel_to_applicant = "08" OR hh_memb_rel_to_applicant = "16" OR hh_memb_rel_to_applicant = "17" Then 
                                                                child_of_hh_member = True
                                                            Else
                                                                child_of_hh_member = False
                                                            End If

                                                            If under_18_check = True and child_of_hh_member = True Then
                                                                If activate_msg_boxes = True then MsgBox "Testing -- under_18_check = True and child_of_hh_member = True. Navigating to SCHL now"
                                                                'Navigate to SCHL panel to check status
                                                                EMWriteScreen "SCHL", 20, 71
                                                                Call write_value_and_transmit(HIRE_memb_number, 20, 76)
                                                                EMReadScreen schl_panel_exists, 25, 24, 2
                                                                If InStr(schl_panel_exists, "DOES NOT EXIST") Then
                                                                    school_status_qualifies = False
                                                                    school_type_qualifies = False
                                                                Else
                                                                    EMReadScreen school_status, 1, 6, 40
                                                                    If school_status = "F" or school_status = "H" Then
                                                                        school_status_qualifies = True
                                                                    Else
                                                                        school_status_qualifies = False
                                                                    End If 

                                                                    EMReadScreen school_type, 2, 7, 40
                                                                    If school_type = "01" or school_type = "11" or school_type = "02" or school_type = "03" Then
                                                                        school_type_qualifies = True
                                                                    Else
                                                                        school_type_qualifies = False
                                                                    End If 

                                                                    EMReadScreen fs_eligibility_status_check, 2, 16, 63
                                                                    If fs_eligibility_status_check = "01" Then 
                                                                        fs_eligibility_eligible = True
                                                                    Else
                                                                        fs_eligibility_eligible = False
                                                                    End If

                                                                End If
                                                            End If

                                                            If under_18_check = True and child_of_hh_member = True and school_status_qualifies = True and school_type_qualifies = True Then
                                                                snap_earned_income_minor_exclusion = True
                                                            Else
                                                                snap_earned_income_minor_exclusion = False
                                                            End If
                                                                
                                                            If snap_earned_income_minor_exclusion = True and fs_eligibility_eligible = True Then
                                                                'Since household member meets exclusion criteria, then HIRE message can just be deleted
                                                                'Navigate to CASE/NOTE
                                                                If activate_msg_boxes = True then msgbox "Testing -- snap_earned_income_minor_exclusion = True. navigating to create CASE/NOTE"
                                                                PF4

                                                                EMReadScreen case_note_check, 4, 2, 45
                                                                If case_note_check <> "NOTE" then MsgBox "Testing -- not at case note stop here5"
                                                                'Open a new case note
                                                                PF9

                                                                CALL write_variable_in_case_note("-NDNH Match for (M" & HIRE_memb_number & ") for " & trim(HIRE_employer_name) & "-")
                                                                CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
                                                                CALL write_variable_in_case_note("MAXIS NAME: " & NDNH_maxis_name)
                                                                CALL write_variable_in_case_note("NEW HIRE NAME: " & NDNH_new_hire_name)
                                                                CALL write_variable_in_case_note("EMPLOYER: " & HIRE_employer_name)
                                                                CALL write_variable_in_case_note("---")
                                                                CALL write_variable_in_case_note("HIRE MESSAGE CLEARED THROUGH INFC. NO JOBS PANEL CREATED. HOUSEHOLD MEMBER APPEARS TO MEET SNAP EARNED INCOME EXCLUSION. SEE CM 0017.15.15 - INCOME OF MINOR CHILD/CAREGIVER UNDER 20.")
                                                                CALL write_variable_in_case_note("---")
                                                                CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING OR THE CM GUIDE TO SIX MONTH BUDGETING.")
                                                                CALL write_variable_in_case_note("---")
                                                                CALL write_variable_in_case_note(worker_signature)

                                                                If activate_msg_boxes = True then MsgBox "Testing -- The script is about to save the CASE/NOTE. Stop here if in testing or production"

                                                                'PF3 to save the CASE/NOTE
                                                                PF3

                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " Household member meets SNAP earned income exclusion. No JOBS panel(s) evaluated or added for member number: " & HIRE_memb_number & ". CASE/NOTE added. Message should be deleted.")

                                                                'PF3 BACK to SCHL panel
                                                                PF3

                                                            ElseIf snap_earned_income_minor_exclusion = True and fs_eligibility_eligible = False Then
                                                                ' MsgBox "Testing -- not 01 on FS eligibility"

                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " HH M" & HIRE_memb_number & " appears to meet SNAP earned income exclusion, however, FS eligibility is not 01 on SCHL panel." & " Message should not be deleted."

                                                            Elseif snap_earned_income_minor_exclusion = False Then
                                                            
                                                                'Navigate to STAT/JOBS to check if corresponding JOBS panel exists
                                                                If activate_msg_boxes = True then msgbox "Testing -- snap_earned_income_minor_exclusion = False so navigating to STAT/JOBS"
                                                                Call write_value_and_transmit("JOBS", 20, 71)

                                                                'Open the first JOBS panel of the caregiver reference number
                                                                EMWriteScreen HIRE_memb_number, 20, 76
                                                                Call write_value_and_transmit("01", 20, 79)

                                                                'Delete after testing
                                                                ' msgbox "Delete after testing -- about to add new JOBS panel 5979"
                                                                Call check_and_add_new_jobs_panel(False)
                                                                
                                                            End If
                                                        End If
                                                    Else
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "The employer name and/or state is blank." & " Message should not be deleted."
                                                        
                                                    End If

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should not be deleted") Then
                                                        'The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        If activate_msg_boxes = True then MsgBox "Testing -- Adding to skip list"
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should be deleted") Then 
                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "A JOBS panel exists for employer") Then
                                                            'There was already a corresponding JOBS panel for the employer. The message needs to be deleted through the INFC as a known job.
                                                            list_of_DAIL_messages_to_delete_NDNH_known = list_of_DAIL_messages_to_delete_NDNH_known & full_dail_msg & "*"
                                                            If activate_msg_boxes = True then msgbox "Testing -- Adding to NDNH known delete list"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            dail_row = dail_row - 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No JOBS panels exist for household member number") Then
                                                            'There were no JOBS panels for the HH Memb so a JOBS panel was created. The message needs to be deleted as an unknown job.
                                                            list_of_DAIL_messages_to_delete_NDNH_not_known = list_of_DAIL_messages_to_delete_NDNH_not_known & full_dail_msg & "*"
                                                            If activate_msg_boxes = True then msgbox "Testing -- Adding to NDNH not known delete list. NOT SNAP EXCLUSION"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            dail_row = dail_row - 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Household member meets SNAP earned income exclusion") Then
                                                            'HH MEMB appears to meet SNAP EARNED INCOME EXLCUSION SO MESSAGE Can just be deleted.
                                                            list_of_DAIL_messages_to_delete_NDNH_not_known = list_of_DAIL_messages_to_delete_NDNH_not_known & full_dail_msg & "*"
                                                            If activate_msg_boxes = True then msgbox "Testing -- SNAP Exclusion. Adding to NDNH not known delete list."
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            dail_row = dail_row - 1
                                                        Else
                                                            msgbox "Testing - There was a messsage that did not meet any criteria after determining it was a delete message. Something went wrong. Line 5280"
                                                        End If
                                                    Else
                                                        msgbox "Testing - There was a messsage that did not meet any criteria. Something went wrong. Line 5283"
                                                    End If

                                                    'PF3 back to DAIL
                                                    PF3

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should be deleted") Then
                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No JOBS panels exist for household member number") Then
                                                            ' EMWaitReady 1, 1000
                                                        End If
                                                    End If
                                                    
                                                    Call nav_back_to_dail_check(True)

                                                    'Navigate back to DAIL message - case name and number
                                                    EMWriteScreen hire_message_member_name, 21, 25
                                                    transmit

                                                ElseIf InStr(dail_msg, "NEW JOB DETAILS FOR SSN") Then
                                                    'No action on these, simply note in spreadsheet that QI team to review

                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = "NEW JOB DETAILS FOR SSN message. Outdated HIRE message."

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                    'Update the excel spreadsheet with processing notes
                                                    objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                ElseIf InStr(dail_msg, "SDNH NEW JOB DETAILS") Then

                                                    'Reset variables to ensure they don't carry forward through do loop
                                                    hire_sdnh_message_standardized = ""
                                                    no_exact_JOBS_panel_matches = ""
                                                    list_of_employers_on_jobs_panels = "*"
                                                    JOBS_footer_month = ""
                                                    JOBS_footer_year = ""
                                                    HIRE_case_number = ""
                                                    HIRE_case_name = ""
                                                    HIRE_memb_number = ""
                                                    date_hired = ""
                                                    HIRE_employer_name = ""
                                                    SDNH_MAXIS_name = ""
                                                    SDNH_new_hire_name = ""
                                                    hire_message_member_name = ""
                                                    hire_message_case_number = ""
                                                    HIRE_employer_name_first_word = ""
                                                    HIRE_employer_name_TIKL = ""
                                                    tikl_case_number = ""
                                                    tikl_case_name = ""

                                                    'Blanking variables to check for potential SNAP income exclusion
                                                    hh_memb_age = ""
                                                    under_18_check = ""
                                                    hh_memb_rel_to_applicant = ""
                                                    child_of_hh_member = ""
                                                    school_status = ""
                                                    school_status_qualifies = ""
                                                    school_type = ""
                                                    school_type_qualifies = ""
                                                    snap_earned_income_minor_exclusion = ""
                                                    fs_eligibility_eligible = ""
                                                    fs_eligibility_status_check = ""

                                                    'Read the HIRE message member name to navigate back if needed
                                                    EMReadScreen hire_message_member_name, 8, dail_row - 1, 5
                                                    EMReadScreen hire_message_case_number, 8, dail_row - 1, 73
                                                    hire_message_case_number = trim(hire_message_case_number)

                                                    'Enters “X” on DAIL message to open full message. 
                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                    'Delete after testing - trying to figure out when and why script sometimes does not clear the X
                                                    EmReadScreen multiple_selections_error_check, 20, 24, 2
                                                    If InStr(multiple_selections_error_check, "YOU MAY ONLY SELECT") Then msgbox "6158 It failed to clear the previous X"

                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    EMReadScreen check_full_dail_msg_case_number, 35, 6, 44
                                                    check_full_dail_msg_case_number = trim(check_full_dail_msg_case_number)
                                                    EMReadScreen check_full_dail_msg_case_name, 35, 7, 44
                                                    check_full_dail_msg_case_name = trim(check_full_dail_msg_case_name)

                                                    EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                    check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                    EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                    check_full_dail_msg_line_2 = trim(check_full_dail_msg_line_2)
                                                    EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                    check_full_dail_msg_line_3 = trim(check_full_dail_msg_line_3)
                                                    EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                    check_full_dail_msg_line_4 = trim(check_full_dail_msg_line_4)

                                                    check_full_dail_msg = trim(check_full_dail_msg_case_number & " " & check_full_dail_msg_case_name & " " & check_full_dail_msg_line_1 & " " & check_full_dail_msg_line_2 & " " & check_full_dail_msg_line_3 & " " & check_full_dail_msg_line_4)

                                                    If check_full_dail_msg <> full_dail_msg Then
                                                        MsgBox "Testing -- messages do not match. check_full_dail_msg " & check_full_dail_msg & "    " & " full_dail_msg " & full_dail_msg
                                                    End if

                                                    'Read the Case Name and Case Number to process TIKLs as needed
                                                    EmReadScreen tikl_case_number, 10, 6, 57
                                                    tikl_case_number = trim(tikl_case_number)
                                                    EmReadScreen tikl_case_name, 25, 7, 55
                                                    tikl_case_name = trim(tikl_case_name)

                                                    'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                    'Set row and col
                                                    row = 1
                                                    col = 1
                                                    EMSearch "Case Number: ", row, col
                                                    EMReadScreen HIRE_case_number, 10, row, col + 13
                                                    HIRE_case_number = trim(HIRE_case_number)

                                                    row = 1
                                                    col = 1
                                                    EMSearch "Case Name: ", row, col
                                                    EMReadScreen HIRE_case_name, 25, row, col + 11
                                                    HIRE_case_name = trim(HIRE_case_name)

                                                    row = 1
                                                    col = 1
                                                    EMSearch "SDNH NEW JOB DETAILS FOR MEMB", row, col
                                                    EMReadScreen HIRE_memb_number, 2, row, col + 30
                                                    HIRE_memb_number = trim(HIRE_memb_number)
                                                    If HIRE_memb_number = "00" then msgbox "Testing -- HH MEMB 00 - this is an error message. Need handling for this situation"

                                                    row = 1
                                                    col = 1
                                                    EMSearch "DATE HIRED:", row, col
                                                    EMReadScreen date_hired, 10, row, col + 12
                                                    date_hired = trim(date_hired)
                                                    'Switch dashes to slashes for consistency with NDNH
                                                    date_hired_NDNH_comparison = replace(date_hired, "-", "/")

                                                    If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then
                                                        msgbox "Testing -- date hired is EM or unknown. How to handle?"
                                                    Else
                                                        Call ONLY_create_MAXIS_friendly_date(date_hired)
                                                        date_split = split(date_hired, "/")
                                                        month_hired = date_split(0)
                                                        day_hired = date_split(1)
                                                        year_hired = date_split(2)
                                                    End if

                                                    row = 1
                                                    col = 1
                                                    EMSearch "EMPLOYER: ", row, col
                                                    'Only capturing first 20 characters to align with NDNH
                                                    EMReadScreen HIRE_employer_name, 20, row, col + 10
                                                    HIRE_employer_name = trim(HIRE_employer_name)
                                                    EMReadScreen HIRE_employer_name_TIKL, 25, row, col + 10
                                                    HIRE_employer_name_TIKL = TRIM(HIRE_employer_name_TIKL)

                                                    'Add handling to compare the first word of employer from HIRE to first word of employer on JOBS panel
                                                    HIRE_employer_name_split = split(HIRE_employer_name, " ")

                                                    If len(HIRE_employer_name_split(0)) < 4 and Ubound(HIRE_employer_name_split) > 0 Then
                                                        HIRE_employer_name_first_word = HIRE_employer_name_split(0) & " " & HIRE_employer_name_split(1)
                                                        If activate_msg_boxes = True then MsgBox "First word less than 3 characters long. HIRE_employer_name_first_word is " & HIRE_employer_name_first_word  
                                                    Else
                                                        HIRE_employer_name_first_word = HIRE_employer_name_split(0)   
                                                        If activate_msg_boxes = True then MsgBox "First word longer than 3 characters long. HIRE_employer_name_first_word is " & HIRE_employer_name_first_word
                                                    End If

                                                    If instr(len(HIRE_employer_name_first_word), HIRE_employer_name_first_word, ",") = len(HIRE_employer_name_first_word) then 
                                                        HIRE_employer_name_first_word = Mid(HIRE_employer_name_first_word, 1, len(HIRE_employer_name_first_word) - 1)
                                                        If activate_msg_boxes = True then MsgBox "Last character is a comma. HIRE_employer_name_first_word is now " & HIRE_employer_name_first_word
                                                    End If

                                                    row = 1
                                                    col = 1
                                                    EMSearch "MAXIS NAME   : ", row, col
                                                    EMReadScreen SDNH_maxis_name, 57, row, col + 15
                                                    SDNH_maxis_name = trim(SDNH_maxis_name)

                                                    row = 1
                                                    col = 1
                                                    EMSearch "NEW HIRE NAME: ", row, col
                                                    EMReadScreen SDNH_new_hire_name, 57, row, col + 15
                                                    SDNH_new_hire_name = trim(SDNH_new_hire_name)

                                                    'Standard NDNH format is *[Case Number]-[Case Name]-[Memb ##]-[Date Hired with slashes (MM/DD/YYYY)]-[Employer - first 20 characters]-[Maxis name]-[new hire name]*
                                                    hire_sdnh_message_standardized = "*" & HIRE_case_number & "-" & HIRE_case_name & "-" & HIRE_memb_number & "-" & date_hired_NDNH_comparison & "-" & HIRE_employer_name & "-" & SDNH_maxis_name & "-" & SDNH_new_hire_name & "*"

                                                    If Instr(list_of_NDNH_messages_standard_format, hire_sdnh_message_standardized) then 
                                                        If activate_msg_boxes = True then MsgBox "Testing -- duplicate SDNH message. It will get added to delete list."
                                                    
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = "Duplicate SDNH message. Message should be deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        
                                                        list_of_DAIL_messages_to_delete_SDNH = list_of_DAIL_messages_to_delete_SDNH & full_dail_msg & "*"
                                                        
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        dail_row = dail_row - 1
                                                        
                                                        'Transmit back to DAIL
                                                        transmit

                                                    Else
                                                        If activate_msg_boxes = True then MsgBox "Testing -- Not a duplicate SDNH. Will transmit and process accordingly."

                                                        'Transmit back to DAIL
                                                        transmit

                                                        EMWriteScreen "S", dail_row, 3
                                                        EMSendKey "<enter>"
                                                        EMReadScreen background_check, 25, 7, 30
                                                        If InStr(background_check, "A Background transaction") Then
                                                            EMWaitReady 2, 2000
                                                            Do
                                                                background_check = ""
                                                                PF3
                                                                EMWaitReady 2, 2000
                                                                EMWriteScreen "S", dail_row, 3
                                                                EMWaitReady 2, 2000
                                                                EMSendKey "<enter>"
                                                                EMWaitReady 2, 2000
                                                                EMReadScreen background_check, 25, 7, 30
                                                                If InStr(background_check, "A Background transaction") = 0 then Exit Do
                                                            Loop
                                                            
                                                        End If

                                                        EMReadScreen self_panel_check, 4, 2, 50
                                                        If self_panel_check = "SELF" Then
                                                            EMWaitReady 2, 2000
                                                            EMWaitReady 2, 2000
                                                            EMWriteScreen "DAIL", 16, 43
                                                            EMWriteScreen "DAIL", 21, 70
                                                            transmit

                                                            EMReadSCreen back_to_dail_check, 8, 1, 72
                                                            If back_to_dail_check = "FMKDLAM6" Then

                                                                'Navigate to CASE/CURR to force DAIL to reset and then PF3 back to get back to start of the DAIL
                                                                Call write_value_and_transmit("H", dail_row, 3)
                                                                PF3

                                                                'Reset DAIL to only HIRE messages
                                                                Call write_value_and_transmit("X", 4, 12)
                                                                EMWriteScreen "_", 7, 39
                                                                Call write_value_and_transmit("X", 13, 39)

                                                                'Script should now navigate to specific member name, or at least get close
                                                                EMWriteScreen hire_message_member_name, 21, 25
                                                                transmit

                                                                'Script will enter do loop to find match

                                                                Do
                                                                    return_full_dail_msg = ""
                                                                    return_full_dail_msg_case_number = ""
                                                                    return_full_dail_msg_case_name = ""
                                                                    return_full_dail_msg_line_1 = ""
                                                                    return_full_dail_msg_line_2 = ""
                                                                    return_full_dail_msg_line_3 = ""
                                                                    return_full_dail_msg_line_4 = ""

                                                                    'Enters “X” on DAIL message to open full message. 
                                                                    Call write_value_and_transmit("X", dail_row, 3)

                                                                    'Delete after testing - trying to figure out when and why script sometimes does not clear the X
                                                                    EmReadScreen multiple_selections_error_check, 20, 24, 2
                                                                    If InStr(multiple_selections_error_check, "YOU MAY ONLY SELECT") Then msgbox "6346 It failed to clear the previous X"

                                                                    ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                                    EMReadScreen return_full_dail_msg_case_number, 35, 6, 44
                                                                    return_full_dail_msg_case_number = trim(return_full_dail_msg_case_number)
                                                                    EMReadScreen return_full_dail_msg_case_name, 35, 7, 44
                                                                    return_full_dail_msg_case_name = trim(return_full_dail_msg_case_name)

                                                                    EMReadScreen return_full_dail_msg_line_1, 60, 9, 5
                                                                    return_full_dail_msg_line_1 = trim(return_full_dail_msg_line_1)
                                                                    EMReadScreen return_full_dail_msg_line_2, 60, 10, 5
                                                                    return_full_dail_msg_line_2 = trim(return_full_dail_msg_line_2)
                                                                    EMReadScreen return_full_dail_msg_line_3, 60, 11, 5
                                                                    return_full_dail_msg_line_3 = trim(return_full_dail_msg_line_3)
                                                                    EMReadScreen return_full_dail_msg_line_4, 60, 12, 5
                                                                    return_full_dail_msg_line_4 = trim(return_full_dail_msg_line_4)

                                                                    return_full_dail_msg = trim(return_full_dail_msg_case_number & " " & return_full_dail_msg_case_name & " " & return_full_dail_msg_line_1 & " " & return_full_dail_msg_line_2 & " " & return_full_dail_msg_line_3 & " " & return_full_dail_msg_line_4)

                                                                    If return_full_dail_msg = check_full_dail_msg Then 
                                                                        transmit
                                                                        Exit Do
                                                                    Else
                                                                        transmit
                                                                        dail_row = dail_row + 1

                                                                        'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                                                                        EMReadScreen new_case, 8, dail_row, 63
                                                                        new_case = trim(new_case)
                                                                        IF new_case <> "CASE NBR" THEN 
                                                                            'If there is NOT a new case number, the script will top the message
                                                                            Call write_value_and_transmit("T", dail_row, 3)
                                                                        ELSEIF new_case = "CASE NBR" THEN
                                                                            'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                                                                            Call write_value_and_transmit("T", dail_row + 1, 3)
                                                                        End if
                                                                    End If

                                                                Loop

                                                                'Reset the dail_row back to 6
                                                                dail_row = 6

                                                                EMWriteScreen "S", dail_row, 3
                                                                EMSendKey "<enter>"
                                                            Else
                                                                'Initial dialog - select whether to create a list or process a list
                                                                Dialog1 = ""
                                                                BeginDialog Dialog1, 0, 0, 306, 220, "Unable to return to DAIL. Double-check the issue."
                                                                
                                                                ButtonGroup ButtonPressed
                                                                    OkButton 205, 200, 40, 15
                                                                    CancelButton 245, 200, 40, 15
                                                                EndDialog

                                                                Do
                                                                    Dialog Dialog1
                                                                Loop until ButtonPressed = OK
                                                                
                                                            End If
                                                        End If

                                                        EMWriteScreen "MEMB", 20, 71
                                                        Call write_value_and_transmit(HIRE_memb_number, 20, 76)

                                                        EMReadScreen memb_panel_check, 4, 2, 48
                                                        IF memb_panel_check <> "MEMB" Then 
                                                            EMReadScreen summ_panel_check, 4, 2, 46
                                                            If summ_panel_check = "SUMM" Then
                                                                EMWriteScreen "MEMB", 20, 71
                                                                Call write_value_and_transmit(HIRE_memb_number, 20, 76)
                                                                EMReadScreen memb_panel_check, 4, 2, 48
                                                                IF memb_panel_check <> "MEMB" Then MsgBox "Testing -- second attempt to get to MEMB failed 5709"
                                                            Else
                                                                MsgBox "Testing -- not on Summ 5709. Will attempt to go back to DAIL"
                                                            End If
                                                        End If

                                                        'Ensure the script is not creating a new MEMB panel
                                                        EMReadScreen new_memb_panel_check, 12, 8, 22
                                                        If new_memb_panel_check = "Arrival Date" Then
                                                            PF3
                                                            PF10
                                                            msgbox "Testing -- Script tried to navigate to a HH Memb that doesn't exist. It should have deleted the panel but double check MAKE SURE IT DELETED ADDED PANEL"
    
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "HIRE message identifies a HH Member that does not exist on the case (" & HIRE_memb_number & "). Review needed." & " Message should not be deleted."
                                                        Else
                                                        
                                                            'Check the HH Memb's age and relationship status
                                                            EMReadScreen hh_memb_age, 2, 8, 76
                                                            hh_memb_age = trim(hh_memb_age)
                                                            'Convert age into a number
                                                            If hh_memb_age = "" then MsgBox "No age on panel. stop here"
                                                            If hh_memb_age <> "" Then hh_memb_age = hh_memb_age * 1

                                                            If hh_memb_age > 17 then 
                                                                under_18_check = False 
                                                            Else
                                                                under_18_check = True
                                                            End If

                                                            'Convert age to a number
                                                            EMReadScreen hh_memb_rel_to_applicant, 2, 10, 42
                                                            If hh_memb_rel_to_applicant = "03" OR hh_memb_rel_to_applicant = "08" OR hh_memb_rel_to_applicant = "16" OR hh_memb_rel_to_applicant = "17" Then 
                                                                child_of_hh_member = True
                                                            Else
                                                                child_of_hh_member = False
                                                            End If

                                                            If under_18_check = True and child_of_hh_member = True Then
                                                                If activate_msg_boxes = True then MsgBox "Testing -- under_18_check = True and child_of_hh_member = True. Navigating to SCHL now"
                                                                'Navigate to SCHL panel to check status
                                                                EMWriteScreen "SCHL", 20, 71
                                                                Call write_value_and_transmit(HIRE_memb_number, 20, 76)
                                                                EMReadScreen schl_panel_exists, 25, 24, 2
                                                                If InStr(schl_panel_exists, "DOES NOT EXIST") Then
                                                                    school_status_qualifies = False
                                                                    school_type_qualifies = False
                                                                Else
                                                                    EMReadScreen school_status, 1, 6, 40
                                                                    If school_status = "F" or school_status = "H" Then
                                                                        school_status_qualifies = True
                                                                    Else
                                                                        school_status_qualifies = False
                                                                    End If 

                                                                    EMReadScreen school_type, 2, 7, 40
                                                                    If school_type = "01" or school_type = "11" or school_type = "02" or school_type = "03" Then
                                                                        school_type_qualifies = True
                                                                    Else
                                                                        school_type_qualifies = False
                                                                    End If 

                                                                    EMReadScreen fs_eligibility_status_check, 2, 16, 63
                                                                    If fs_eligibility_status_check = "01" Then 
                                                                        fs_eligibility_eligible = True
                                                                    Else
                                                                        fs_eligibility_eligible = False
                                                                    End If
                                                                End If
                                                            End If

                                                            If under_18_check = True and child_of_hh_member = True and school_status_qualifies = True and school_type_qualifies = True Then
                                                                snap_earned_income_minor_exclusion = True
                                                            Else
                                                                snap_earned_income_minor_exclusion = False
                                                            End If
                                                                
                                                            If snap_earned_income_minor_exclusion = True and fs_eligibility_eligible = True Then
                                                                'Since household member meets exclusion criteria, then HIRE message can just be deleted
                                                                If activate_msg_boxes = True then MsgBox "Testing -- Navigating to CASE/NOTE. stop here if needed"
                                                                
                                                                'Navigate to CASE/NOTE
                                                                PF4
                                                                EMReadScreen case_note_check, 4, 2, 45
                                                                If case_note_check <> "NOTE" then MsgBox "Testing -- not at case note stop here6"

                                                                'Open a new case note
                                                                PF9

                                                                CALL write_variable_in_case_note("-SDNH Match for (M" & HIRE_memb_number & ") for " & trim(HIRE_employer_name) & "-")
                                                                CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
                                                                CALL write_variable_in_case_note("MAXIS NAME: " & SDNH_maxis_name)
                                                                CALL write_variable_in_case_note("NEW HIRE NAME: " & SDNH_new_hire_name)
                                                                CALL write_variable_in_case_note("EMPLOYER: " & HIRE_employer_name)
                                                                CALL write_variable_in_case_note("---")
                                                                CALL write_variable_in_case_note("HIRE MESSAGE DELETED. NO JOBS PANEL CREATED. HOUSEHOLD MEMBER APPEARS TO MEET SNAP EARNED INCOME EXCLUSION. SEE CM 0017.15.15 - INCOME OF MINOR CHILD/CAREGIVER UNDER 20.")
                                                                CALL write_variable_in_case_note("---")
                                                                CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A 6-MONTH REPORTING CASE. SEE 0007.03.02 - SIX-MONTH REPORTING.")
                                                                CALL write_variable_in_case_note("---")
                                                                CALL write_variable_in_case_note(worker_signature)

                                                                If activate_msg_boxes = True then MsgBox "Testing -- The script is about to save the CASE/NOTE re: SNAP exclusion. Stop here if in testing or production"

                                                                'PF3 to save the CASE/NOTE
                                                                PF3

                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " Household member meets SNAP earned income exclusion. No JOBS panel(s) evaluated or added for member number: " & HIRE_memb_number & ". CASE/NOTE added. Message should be deleted.")

                                                                'PF3 BACK to SCHL panel
                                                                PF3

                                                            ElseIf snap_earned_income_minor_exclusion = True and fs_eligibility_eligible = False Then
                                                                ' MsgBox "Testing -- not 01 on FS eligibility"

                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " HH M" & HIRE_memb_number & " appears to meet SNAP earned income exclusion, however, FS eligibility is not 01 on SCHL panel." & " Message should not be deleted."


                                                            Elseif snap_earned_income_minor_exclusion = False Then
                                                                
                                                                If activate_msg_boxes = True then MsgBox "Testing -- Not snap income exclusion. Navigate to JOBS."
                                                                
                                                                'Navigate to STAT/JOBS to check if corresponding JOBS panel exists
                                                                Call write_value_and_transmit("JOBS", 20, 71)

                                                                EMReadScreen jobs_panel_nav_check, 8, 2, 43
                                                                If InStr(jobs_panel_nav_check, "JOBS") = 0 Then MsgBox "Testing -- Stop here. Not at JOBS panel"

                                                                'Open the first JOBS panel of the HH memb number
                                                                EMWriteScreen HIRE_memb_number, 20, 76
                                                                Call write_value_and_transmit("01", 20, 79)

                                                                'Delete after testing
                                                                ' msgbox "Delete after testing -- about to add new JOBS panel 6477"
                                                                Call check_and_add_new_jobs_panel(False)
                                                                
                                                            End If
                                                        End If
                                                
                                                        If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should not be deleted") Then
                                                            If activate_msg_boxes = True then msgbox "Testing -- add to skip list for SDNH"
                                                            'The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                            list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                        ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should be deleted") Then
                                                            If activate_msg_boxes = True then msgbox "Testing -- add to delete list for SDNH"
                                                            'There is a corresponding JOBS panel or a JOBS panel was created. The message can be deleted.
                                                            list_of_DAIL_messages_to_delete_SDNH = list_of_DAIL_messages_to_delete_SDNH & full_dail_msg & "*"
                                                            'Update the excel spreadsheet with processing notes
                                                            objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                            dail_row = dail_row - 1
                                                        Else
                                                            MsgBox "Testing -- Script reached a SDNH message that does not meet delete or not delete criteria."
                                                        End If

                                                        'PF3 back to DAIL
                                                        PF3
                                                        
                                                    End If

                                                    Call nav_back_to_dail_check(True)
                                                    
                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should be deleted") Then 
                                                        ' EMWaitReady 1, 1000
                                                    End If

                                                    'Navigate back to DAIL message - case name and number
                                                    EMWriteScreen hire_message_member_name, 21, 25
                                                    transmit

                                                ElseIf InStr(dail_msg, "JOB DETAILS FOR  ") Then
                                                    'No action on these, simply note in spreadsheet that QI team to review

                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = "QI Review. Outdated HIRE message."

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                    'Update the excel spreadsheet with processing notes
                                                    objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                Else
                                                    ' Add handling as needed here
                                                End If
                                            Else
                                                ' Add handling as needed here
                                            End If

                                        End If

                                        'Increment the dail_excel_row so that data isn't overwritten
                                        dail_excel_row = dail_excel_row + 1
                                        
                                        'Increment dail_count for the dail array
                                        dail_count = dail_count + 1

                                        'In instances where the case details are not the final item in the array, need to exit the for loop
                                        Exit For

                                    End If 
                                Next

                            Else
                                'Add handling for messages that are not meeting any criteria. May not be necessary but have this just in case
                                msgbox "Testing -- Instance where it is NOT on the delete list, not on the skip list, and not on either list. So could be a repeat or something?"
                            End If
                                
                        End If
                    Else
                        'Potentially add handling for cases that are not on valid case numbers list, just set processable to false and include processing note that it is likely out of county or privileged?
                    
                    End If
                            
                
                Else
                    'Add handling as needed
                End If

                ' 'Increment the stats counter
                stats_counter = stats_counter + 1
                
                dail_row = dail_row + 1

                'Checking for the last DAIL message. If it just processed the final message, the DAIL will appear blank but there is actually an invisible '_' at 6, 3. Handling to check for this and then navigate to the next page if needed. If it is on the last page, then it will exit the do loop 
                EMReadScreen next_dail_check, 7, dail_row, 3
                If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    'Attempt to navigate to the next page
                    PF8
                    EMReadScreen last_page_check, 21, 24, 2
                    'Check if the last page of the DAIL has been reached, also handles for situations where the last DAIL has been deleted and it displays a 'NO MESSAGES' warning
                    If last_page_check = "THIS IS THE LAST PAGE" or Instr(last_page_check, "NO MESSAGES") then
                        all_done = true
                        exit do
                    Else
                        dail_row = 6
                    End if
                End if
            LOOP
            IF all_done = true THEN exit do
        LOOP

        'Now that the script has compiled a string of all of the NDNH messages, it will now evaluate the individual messages to determine if there is a duplicate SDNH, or if it can process the SDNH or NDNH message
        'Reset the all_done so that it doesn't exit after the first run unintentionally
        all_done = ""

        If activate_msg_boxes = True then MsgBox "Testing -- script successfully processed HIRE messages. It will now review TIKLs. list_of_TIKLs_to_delete " & list_of_TIKLs_to_delete

        'Navigate to TIKLs for the X number
        'Set the TIKLs to first of next month

        EmWriteScreen CM_plus_2_mo, 4, 67
        EmWriteScreen "01", 4, 70
        EmWriteScreen CM_plus_2_yr, 4, 73
        Call write_value_and_transmit("X", 4, 12)
        EmWriteScreen "_", 7, 39
        EmWriteScreen "X", 19, 39
        transmit

        'The script should be back at start of TIKLs for correct month
        'Reads where the count of DAILs is listed. Used to verify DAIL is not empty.
        EMReadScreen number_of_dails, 1, 3, 67		

        DO
            'If this space is blank the rest of the DAIL reading is skipped
            If number_of_dails = " " Then exit do		
            'Because the script brings each new case to the top of the page, dail_row starts at 6.
            dail_row = 6	

            DO

                tikl_case_name_check = ""
                tikl_case_number_check = ""
                
                'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
                EMReadScreen new_case, 8, dail_row, 63
                new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN 
                    'If there is NOT a new case number, the script will top the message
                    Call write_value_and_transmit("T", dail_row, 3)
                ELSEIF new_case = "CASE NBR" THEN
                    'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                End if

                'Resets the DAIL row since the message has now been topped
                dail_row = 6  

                'Determines the DAIL Type
                EMReadScreen dail_type, 4, dail_row, 6
                dail_type = trim(dail_type)

                'Determines the TIKL date
                EMReadScreen tikl_date, 8, dail_row, 11
                tikl_date = trim(tikl_date)

                'Reads the DAIL msg to determine if it is an out-of-scope message
                EMReadScreen dail_msg, 61, dail_row, 20
                dail_msg = trim(dail_msg)

                EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                MAXIS_case_number = trim(MAXIS_case_number)

                If (InStr(dail_msg, "VERIFICATION OF ") <> 0 and Instr(dail_msg, "VIA NEW HIRE") <> 0 and Instr(dail_msg, " JOB (HIRE") = 0) or (InStr(dail_msg, "VERIFICATION OF ") <> 0 and Instr(dail_msg, " JOB (HIRE") <> 0) Then
                    'The DAIL message is a TIKL for new hire script

                    Call write_value_and_transmit("X", dail_row, 3)

                    'Read the Case Name and Case Number to process TIKLs as needed
                    EmReadScreen tikl_case_number_check, 10, 6, 57
                    tikl_case_number_check = trim(tikl_case_number_check)
                    EmReadScreen tikl_case_name_check, 25, 7, 55
                    tikl_case_name_check = trim(tikl_case_name_check)
                    transmit

                    If InStr(dail_msg, "VERIFICATION OF ") <> 0 and Instr(dail_msg, "VIA NEW HIRE") <> 0 and Instr(dail_msg, " JOB (HIRE") = 0 Then
                        TIKL_comparison = "*" & tikl_case_number_check & "-" & tikl_case_name_check & "-" & Mid(dail_msg, 1, instr(dail_msg, "JOB VIA NEW") - 1) & "*"
                        ' If activate_msg_boxes = True then msgbox "TIKL_comparison " & TIKL_comparison & " and the dail_msg is " & dail_msg
                        ' Msgbox "TIKL_comparison " & TIKL_comparison & " and the dail_msg is " & dail_msg

                        
                    ElseIf InStr(dail_msg, "VERIFICATION OF ") <> 0 and Instr(dail_msg, " JOB (HIRE") <> 0 Then
                        TIKL_comparison = "*" & tikl_case_number_check & "-" & tikl_case_name_check & "-" & Mid(dail_msg, 1, instr(dail_msg, " JOB (HIRE") - 1) & "*"
                        ' If activate_msg_boxes = True then msgbox "TIKL_comparison " & TIKL_comparison & " and the dail_msg is " & dail_msg
                        ' Msgbox "TIKL_comparison " & TIKL_comparison & " and the dail_msg is " & dail_msg

                    Else
                        MsgBox "Neither TIKL Worked 5708"
                    End If

                    

                    If InStr(list_of_TIKLs_to_delete, TIKL_comparison) Then 
                        'This is a match for the TIKL, it can be deleted
                        If activate_msg_boxes = True then msgbox "Testing -- found a TIKL match!"
                        'Activate the case details sheet
                        objExcel.Worksheets("HIRE TIKLs").Activate

                        'Add details for tracking TIKLs
                        objExcel.Cells(TIKL_excel_row, 1).Value = tikl_case_number_check
                        objExcel.Cells(TIKL_excel_row, 2).Value = tikl_case_name_check 
                        objExcel.Cells(TIKL_excel_row, 3).Value = dail_type 
                        objExcel.Cells(TIKL_excel_row, 4).Value = tikl_date 
                        objExcel.Cells(TIKL_excel_row, 5).Value = dail_msg 
                        objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL match found. Should be deleted." 
                        
                        'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
                        EMReadScreen last_dail_check, 12, 3, 67
                        last_dail_check = trim(last_dail_check)

                        'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
                        last_dail_check = split(last_dail_check, " ")

                        If last_dail_check(0) = last_dail_check(2) then 
                            'The script is about to delete the LAST message in the DAIL so script will exit do loop after deletion, also works if it is about to delete the ONLY message in the DAIL
                            all_done = true
                        End If

                        If activate_msg_boxes = True then msgbox "Testing -- make sure it added initial info to excel sheet correctly. It is about to delete the TIKL message. Confirm before proceeding."

                        'Delete the message
                        Call write_value_and_transmit("D", dail_row, 3)

                        'Handling for deleting message under someone else's x number
                        EMReadScreen other_worker_error, 25, 24, 2
                        other_worker_error = trim(other_worker_error)

                        If other_worker_error = "ALL MESSAGES WERE DELETED" Then
                            'Script deleted the final message in the DAIL
                            dail_row = dail_row - 1
                            objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message successfully deleted."
                            'Exit do loop as all messages are deleted
                            all_done = true

                        ElseIf other_worker_error = "" Then
                            'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion

                            'Handling to check if message actually deleted
                            total_dail_msg_count_before = last_dail_check(2) * 1
                            EMReadScreen total_dail_msg_count_after, 12, 3, 67

                            total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                            total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                            If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                'The total DAILs decreased by 1, message deleted successfully
                                dail_row = dail_row - 1
                                objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message successfully deleted."
                            Else
                                'The total DAILs did not decrease by 1, something went wrong
                                objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message unable to be deleted for some reason."
                                script_end_procedure_with_error_report("Script end error - something went wrong with deleting the TIKL message 6854.")
                            End If

                        ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 

                            If activate_msg_boxes = True then MsgBox "Testing -- It will transmit again to delete the TIKL"
                            
                            'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
                            transmit

                            'Reads the total number of DAILS after deleting to determine if it decreased by 1
                            EMReadScreen total_dail_msg_count_after, 12, 3, 67

                            'Checks if final DAIL message deleted
                            EMReadScreen final_dail_error, 25, 24, 2

                            If final_dail_error = "ALL MESSAGES WERE DELETED" Then
                                'All DAIL messages deleted so indicates deletion a success
                                dail_row = dail_row - 1
                                objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message successfully deleted."
                                'No more DAIL messages so exit do loop
                                all_done = True
                            ElseIf trim(final_dail_error) = "" Then
                                'Handling to check if message actually deleted
                                total_dail_msg_count_before = last_dail_check(2) * 1

                                total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                    'The total DAILs decreased by 1, message deleted successfully
                                    dail_row = dail_row - 1
                                    objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message successfully deleted."
                                Else
                                    objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message unable to be deleted for some reason."
                                    script_end_procedure_with_error_report("Script end error - something went wrong with deleting the TIKL message 6887.")
                                End If

                            Else
                                objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message unable to be deleted for some reason."
                                script_end_procedure_with_error_report("Script end error - something went wrong with deleting the TIKL message 6892.")
                            End if
                            
                        Else
                            objExcel.Cells(TIKL_excel_row, 6).Value = "TIKL message unable to be deleted for some reason."
                            script_end_procedure_with_error_report("Script end error - something went wrong with deleting the TIKL message - 6897.")
                        End If
                        
                        If activate_msg_boxes = True then msgbox "Testing -- make sure it updated excel sheet correctly"
                        TIKL_excel_row = TIKL_excel_row + 1
                    
                    Else
                        If activate_msg_boxes = True then MsgBox "No match found 6912"

                    End If
                End If
                        
                dail_row = dail_row + 1

                'Checking for the last DAIL message. If it just processed the final message, the DAIL will appear blank but there is actually an invisible '_' at 6, 3. Handling to check for this and then navigate to the next page if needed. If it is on the last page, then it will exit the do loop 
                EMReadScreen next_dail_check, 7, dail_row, 3
                If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    'Attempt to navigate to the next page
                    PF8
                    EMReadScreen last_page_check, 21, 24, 2
                    'Check if the last page of the DAIL has been reached, also handles for situations where the last DAIL has been deleted and it displays a 'NO MESSAGES' warning
                    If last_page_check = "THIS IS THE LAST PAGE" or Instr(last_page_check, "NO MESSAGES") then
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

    'Update Stats Info

    'Calculate the manual time savings
    total_cases_evaluated = case_excel_row - 2
    STATS_manualtime = (total_cases_evaluated * 30) + (dail_msg_deleted_count * 300) + (not_processable_msg_count * 15) + (QI_flagged_msg_count * 30)

    'Activate the stats sheet
    objExcel.Worksheets("Stats").Activate
    objExcel.Cells(1, 2).Value = case_excel_row - 2
    objExcel.Cells(2, 2).Value = dail_excel_row - 2
    objExcel.Cells(3, 2).Value = not_processable_msg_count
    objExcel.Cells(4, 2).Value = dail_msg_deleted_count
    objExcel.Cells(5, 2).Value = QI_flagged_msg_count
    objExcel.Cells(6, 2).Value = timer - start_time
    objExcel.Cells(7, 2).Value = ((STATS_manualtime) - (timer - start_time)) / 60

    'Finding the right folder to automatically save the file
    this_month = CM_mo & " " & CM_yr
    month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
    unclear_info_folder = replace(this_month, " ", "-") & " DAIL Unclear Info"
    report_date = replace(date, "/", "-")

    'saving the Excel file
    file_info = month_folder & "\" & unclear_info_folder & "\" & report_date & " Unclear Info" & " " & "HIRE" & " " & dail_msg_deleted_count

    'Saves and closes the most recent Excel workbook with the Task based cases to process.
    objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"
    objExcel.ActiveWorkbook.Close
    objExcel.Application.Quit
    objExcel.Quit

    script_end_procedure_with_error_report("Success! Please review the list created for accuracy.")
    
End If

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/03/2024
'--Tab orders reviewed & confirmed----------------------------------------------05/03/2024
'--Mandatory fields all present & Reviewed--------------------------------------05/03/2024
'--All variables in dialog match mandatory fields-------------------------------05/03/2024
'Review dialog names for content and content fit in dialog----------------------05/03/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/03/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------05/03/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/03/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------05/03/2024
'--PRIV Case handling reviewed -------------------------------------------------05/03/2024
'--Out-of-County handling reviewed----------------------------------------------05/03/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/03/2024
'--BULK - review output of statistics and run time/count (if applicable)--------08/06/2024
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/03/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/06/2024
'--Incrementors reviewed (if necessary)-----------------------------------------08/06/2024
'--Denomination reviewed -------------------------------------------------------08/06/2024
'--Script name reviewed---------------------------------------------------------08/06/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/06/2024

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------08/06/2024
'--comment Code-----------------------------------------------------------------08/06/2024
'--Update Changelog for release/update------------------------------------------08/06/2024
'--Remove testing message boxes-------------------------------------------------05/22/2024
'--Remove testing code/unnecessary code-----------------------------------------05/22/2024
'--Review/update SharePoint instructions----------------------------------------08/06/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------08/06/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------08/06/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------08/06/2024
'--Complete misc. documentation (if applicable)---------------------------------08/06/2024
'--Update project team/issue contact (if applicable)----------------------------08/06/2024