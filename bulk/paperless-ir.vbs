'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - PAPERLESS Review.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "60"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("05/14/2018", "Updated the TIKL functionality to write TIKL for the current day of the month.", "Ilse Ferris, Hennepin County")
call changelog_update("12/05/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog paperless_IR_dialog, 0, 0, 241, 130, "PAPERLESS IR"
  EditBox 70, 10, 50, 15, worker_number
  EditBox 190, 10, 15, 15, MAXIS_footer_month
  EditBox 210, 10, 15, 15, MAXIS_footer_year
  CheckBox 10, 30, 145, 10, "Check here to run for the entire agency", whole_county_check
  ButtonGroup ButtonPressed
    OkButton 120, 110, 50, 15
    CancelButton 175, 110, 50, 15
  Text 5, 15, 60, 10, "Worker number(s):"
  Text 125, 15, 65, 10, "Footer month/year:"
  GroupBox 5, 50, 220, 55, "About the Paperless IR script:"
  Text 10, 65, 205, 35, "This script will update REVW for each starred IR, after checking JOBS/BUSI/RBIC for discrepancies. It skips cases that are also reviewing for SNAP. You will have to manually check ELIG/HC for each case and approve the results/case note."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

'establishing variable for the script since most users are approving CM + 1
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'defaulting the cleared date in REVW to the frst of the current month
current_month = CM_mo
current_day = "01"
current_year = CM_yr

EMReadScreen on_revw_panel, 4, 2, 52
If on_revw_panel = "REVW" Then
    EMReadScreen basket_number, 7, 21, 6
    worker_number = trim(basket_number)
End If

DO
	DO
		err_msg = ""
		Dialog paperless_IR_dialog
		If buttonpressed = 0 then stopscript
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If trim(worker_number) <> "" AND Len(worker_number) <> 7 then err_msg = err_msg & vbNewLine & "* You must enter a valid 7 DIGIT worker number."
        If trim(worker_number) = "" AND whole_county_check = unchecked then err_msg = err_msg & vbNewLine & "* You must either list a 7 DIGIT worker number OR indicate the script should be run for the entire county."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If whole_county_check = checked Then
    all_case_numbers_array = " "					'Creating blank variable for the future array
    get_county_code	'Determines worker county code

    call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
    worker_array = Array()
    worker_array = Array(worker_number)
End If
' 
' For each worker in worker_array
'     MsgBox worker
' Next

For each worker in worker_array

    Call back_to_SELF
    Call MAXIS_footer_month_confirmation
    Call navigate_to_MAXIS_screen("rept", "revw")
    EMWriteScreen worker, 21, 6
    EMWriteScreen MAXIS_footer_month, 20, 55
    EMWriteScreen MAXIS_footer_year, 20, 58
    transmit

    EMReadScreen REVW_check, 4, 2, 52
    If REVW_check <> "REVW" then script_end_procedure("You must start this script at the beginning of REPT/REVW. Navigate to the screen and try again!")

    row = 7
    Do
        If row = 19 then
            PF8
            row = 7
            EMReadScreen MAXIS_check, 5, 1, 39
            If MAXIS_check <> "MAXIS" then stopscript
            EMReadScreen last_page_check, 4, 24, 14
        End if
        EMReadScreen MAXIS_case_number, 8, row, 6
        EMReadScreen paperless_check, 1, row, 51
        if paperless_check = "*" then
        	case_number_array = trim(case_number_array & " " & trim(MAXIS_case_number))
        	Stats_counter = Stats_counter + 1
    	End if
        row = row + 1
    Loop until last_page_check = "LAST" or trim(MAXIS_case_number) = ""
Next

case_number_array = split(case_number_array)

For each MAXIS_case_number in case_number_array
    actually_paperless = "" 'Resetting the variable.
    call navigate_to_MAXIS_screen ("stat", "memb")
    call navigate_to_MAXIS_screen ("stat", "jobs")
    EMWriteScreen "01", 20, 76
    transmit
    Do
    	EMReadScreen panel_check, 8, 2, 72
    	current_panel = trim(left(panel_check, 2))
    	total_panels = trim(right(panel_check, 2))
    	EMReadScreen date_check, 8, 9, 49
    	If total_panels <> "0" & date_check = "__ __ __" then actually_paperless = False
    	if current_panel <> total_panels then transmit
    Loop until current_panel = total_panels

    call navigate_to_MAXIS_screen ("stat", "busi")
    EMWriteScreen "01", 20, 76
    transmit
  	Do
    	current_panel = trim(left(panel_check, 2))
		EMReadScreen panel_check, 8, 2, 72
    	total_panels = trim(right(panel_check, 2))
    	EMReadScreen date_check, 8, 5, 71
    	If total_panels <> "0" & date_check = "__ __ __" then actually_paperless = False
    	if current_panel <> total_panels then transmit
  	Loop until current_panel = total_panels

  	call navigate_to_MAXIS_screen ("stat", "rbic")
  	EMWriteScreen "01", 20, 76
  	transmit
  	Do
  		EMReadScreen panel_check, 8, 2, 72
  		current_panel = trim(left(panel_check, 2))
  		total_panels = trim(right(panel_check, 2))
  		EMReadScreen date_check, 8, 6, 68
  		If total_panels <> "0" & date_check = "__ __ __" then actually_paperless = False
  		if current_panel <> total_panels then transmit
  	Loop until current_panel = total_panels

  	If actually_paperless <> False then actually_paperless = True

  	If actually_paperless = True then
    	call navigate_to_MAXIS_screen ("stat", "revw")
    	EMReadScreen SNAP_review_check, 1, 7, 60
    	If SNAP_review_check <> "N" then
			STATS_counter = STATS_counter + 1
	  		cases_to_tikl = cases_to_tikl & "~" & MAXIS_case_number
      		PF9
      		EMWriteScreen "x", 5, 71
      		transmit
      		EMReadScreen renewal_year, 2, 8, 33
      		If renewal_year = "__" then
        		EMReadScreen renewal_year, 2, 8, 77
        		renewal_year_col = 77
      		Else
        		renewal_year_col = 33
      		End if
      		EMWriteScreen left(current_month, 2), 6, 27
      		EMWriteScreen current_day, 6, 30
      		EMWriteScreen right(current_year, 2), 6, 33
      		new_renewal_year = cint(right(current_year, 2)) + 1
      		If current_month = 12 then new_renewal_year = new_renewal_year + 1 'Because otherwise the renewal year will be the current footer month.
      		EMWriteScreen new_renewal_year, 8, renewal_year_col
      		EMWriteScreen "U", 13, 43
      		EMReadScreen spouse_check, 1, 14, 43
      		If spouse_check = "N" then
				PF10
				transmit
			End if
		End if
    End if
Next

If cases_to_tikl <> "" Then
	cases_to_tikl = right(cases_to_tikl, len(cases_to_tikl)-1)
	cases_to_tikl_array = split(cases_to_tikl, "~")
Else
    script_end_procedure("No Paperless IR cases found for the worker.")
End If

For each MAXIS_case_number in cases_to_tikl_array
	navigate_to_MAXIS_screen "DAIL", "WRIT"
    call create_MAXIS_friendly_date(date, 0, 5, 18)
	EMWritescreen "%^% Sent through background using bulk script %^%", 9, 3
	transmit
	EMReadScreen tikl_success, 4, 24, 2
	If tikl_success <> "    " Then MsgBox "This case - " & MAXIS_case_number & " failed to have a TIKL set, track and case note manually"
	PF3
Next

transmit
Do
	PF3
	EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

Stats_counter = Stats_counter - 1
script_end_procedure("Success! All starred (*) IRs have been sent into background, except those with current JOBS/BUSI/RBIC, those who have members other than 01 open, or those who also have SNAP up for review. You must go through and approve these results when they come through background.")
