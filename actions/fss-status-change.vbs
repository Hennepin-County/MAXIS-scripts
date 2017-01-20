'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - FSS STATUS CHANGE.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 270                	'manual run time in seconds
STATS_denomination = "I"       		'I is for Item
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS ===================================================================================================================
BeginDialog fss_status_dialog, 0, 0, 221, 265, "FSS Status Update"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 135, 10, 75, 10, "Reload Client Name", get_client_name_button
  EditBox 30, 25, 25, 15, ref_number
  EditBox 60, 25, 155, 15, client_name
  Text 5, 45, 185, 10, "Select all the ES Status Codes that applies to this client."
  CheckBox 5, 60, 155, 10, "Ill/Incapacitated for more than 60 Days - 23", ill_incap_checkbox
  CheckBox 5, 75, 155, 10, "Care of Ill/Incap Family Member - 24", care_of_ill_Incap_checkbox
  CheckBox 5, 90, 155, 10, "Care of Child Under 12 Months - 25", child_under_one_checkbox
  CheckBox 5, 105, 155, 10, "Family Violence Waiver - 26", fam_violence_checkbox
  CheckBox 5, 120, 155, 10, "Special Medical Criteria - 27", Special_medical_checkbox
  CheckBox 5, 135, 155, 10, "IQ Tested - 28", iq_test_checkbox
  CheckBox 5, 150, 155, 10, "Learning Disabled - 29", learning_disabled_checkbox
  CheckBox 5, 165, 155, 10, "Mentally Ill - 30", mentally_ill_checkbox
  CheckBox 5, 180, 155, 10, "Developmentally Delayed - 31", dev_delayed_checkbox
  CheckBox 5, 195, 155, 10, "Unemployable - 32", unemployable_checkbox
  CheckBox 5, 210, 155, 10, "SSI/RSDI Pending - 33", ssi_pending_checkbox
  CheckBox 5, 225, 155, 10, "Newly Arrived Immigrant - 34", new_imig_checkbox
  ButtonGroup ButtonPressed
    OkButton 110, 245, 50, 15
    CancelButton 165, 245, 50, 15
  Text 5, 30, 25, 10, "Client"
  Text 5, 10, 45, 10, "Case Number"
EndDialog


BeginDialog FSS_final_dialog, 0, 0, 421, 180, "FSS Case Note Information"
  EditBox 65, 5, 350, 15, fss_category_list
  CheckBox 10, 25, 135, 10, "Check here if MFIP Results approved", results_approved_checkbox
  CheckBox 10, 35, 150, 10, "Check here if MFIP Results NOT approved", not_approved_checkbox
  EditBox 90, 50, 325, 15, notes_not_approved
  EditBox 65, 70, 350, 15, other_notes
  EditBox 10, 100, 395, 15, MFIP_results
  ButtonGroup ButtonPressed
    PushButton 10, 120, 75, 15, "Send case to BGTX", CASE_BGTX_button
    PushButton 185, 35, 25, 10, "MEMB", MEMB_button
    PushButton 210, 35, 25, 10, "MEMI", MEMI_button
    PushButton 235, 35, 25, 10, "EMPS", EMPS_button
    PushButton 260, 35, 25, 10, "REVW", REVW_button
    PushButton 285, 35, 25, 10, "MONT", MONT_button
    PushButton 310, 35, 25, 10, "PBEN", PBEN_button
    PushButton 335, 35, 25, 10, "DISA", DISA_button
    PushButton 360, 35, 25, 10, "IMIG", IMIG_button
    PushButton 385, 35, 25, 10, "TIME", TIME_button
  EditBox 230, 160, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 310, 160, 50, 15
    CancelButton 365, 160, 50, 15
  Text 10, 140, 295, 10, "If the case is ready for approval with these results. APP the results before pressing 'OK'."
  Text 10, 75, 45, 10, "Other Notes:"
  GroupBox 5, 90, 410, 65, "MFIP Results"
  Text 165, 165, 60, 10, "Worker Signature:"
  Text 10, 10, 50, 10, "FSS Category:"
  Text 10, 55, 75, 10, "Reason not approved:"
  GroupBox 180, 25, 235, 25, "STAT Navigation:"
  Text 185, 125, 175, 10, "Initial Footer Month/Year of MFIP package to approve"
  EditBox 360, 120, 20, 15, month_to_start
  EditBox 385, 120, 20, 15, year_to_start
EndDialog

'===========================================================================================================================

'FUNCTIONS==================================================================================================================
FUNCTION month_change(interval, starting_month, starting_year, result_month, result_year)
	result_month = abs(starting_month)
	result_year = abs(starting_year)
	valid_month = FALSE
	IF result_month = 1 OR result_month = 2 OR result_month = 3 OR result_month = 4 OR result_month = 5 OR result_month = 6 OR result_month = 7 OR result_month = 8 OR result_month = 9 OR result_month = 10 OR result_month = 11 OR result_month = 12 Then valid_month = TRUE
	If valid_month = FALSE Then
		Month_Input_Error_Msg = MsgBox("The month to start from is not a number between 1 and 12, these are the only valid entries for this function. Your data will have the wrong month." & vbnewline & "The month input was: " & result_month & vbnewline & vbnewline & "Do you wish to continue?", vbYesNo + vbSystemModal, "Input Error")
		If Month_Input_Error_Msg = VBNo Then script_end_procedure("")
	End If
	Do
		If left(interval, 1) = "-" Then
			result_month = result_month - 1
			If result_month = 0 then
				result_month = 12
				result_year = result_year - 1
			End If
			interval = interval + 1
		Else
			result_month = result_month + 1
			If result_month = 13 then
				result_month = 1
				result_year = result_year + 1
			End if
			interval = interval - 1
		End If
	Loop until interval = 0
	result_month = right("00" & result_month, 2)
	result_year = right(result_year, 2)
END FUNCTION

FUNCTION Read_MFIP_Results(initial_month, initial_year, MFIP_results)
	'date_array = null
	'Call date_array_generator(initial_month, initial_year, months_of_mfip_array)

		'THIS IS THE DATE ARRAY GENERATOR - IT WAS CAUSING PROBLEMS TO BE CALLED TWICE================
		'So I embedded the code into this function
		date_list = ""
		'defines an intial date from the initial_month and initial_year parameters
		initial_date = initial_month & "/1/" & initial_year
		'defines a date_list, which starts with just the initial date
		date_list = initial_date
		'This loop creates a list of dates
		Do
			If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
			working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
			date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
		Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

		'Splits this into an array
		months_of_mfip_array = split(date_list, "|")
		'=============================================================================================

	MFIP_results = ""		'Since this is stored as a string, blanking it out so that it doesn't keep old data

	For Each version in months_of_mfip_array
		MAXIS_footer_month = right("00" & datepart("m", version), 2)		'Setting the footer month and year
		MAXIS_footer_year = right(datepart("yyyy", version), 2)
		Back_to_SELF														'Footer month and year do not change well within ELIG
		Call Navigate_to_MAXIS_screen ("ELIG", "MFIP")
		EMReadScreen elig_check, 4, 3, 47									'Makes sure there is an ELIG version to read
		If elig_check = "MFPR" Then
			EMReadScreen process_date, 8, 2, 73								'Makes sure the elig results are from today
			If CDate(process_date) = date Then
				EMWriteScreen "MFSM", 20, 71								'Goes to the last page of ELIG
				transmit
				Do
					EMReadScreen benefit_status, 13, 10, 31					'Sometimes if there is 'NO CHANGE' the benefits incorrectly list as $0
					benefit_status = trim(benefit_status)
					If benefit_status = "NO CHANGE" Then
						no_change = TRUE
						EMReadScreen total_grant, 8, 13, 73
						If trim(total_grant) = "0.00" Then 					'Switches to a previous version that will list the amounts
							EMReadScreen version, 1, 2, 12
							version = abs(version)
							prev_version = version - 1
							EMWriteScreen "0" & prev_version, 20, 79
							transmit
						Else Exit Do
						End If
					End If
				Loop Until benefit_status <> "NO CHANGE"
				EMReadScreen total_grant, 8, 13, 73							'Reads all of the benefit amounts by category
				EMReadScreen cash_amt, 8, 14, 73
				EMReadScreen food_amt, 8, 15, 73
				EMReadScreen housing_grant, 8, 16, 73						'Lists them as a string, formatted with ; for case note display
				MFIP_results = MFIP_results & MAXIS_footer_month & "/" & MAXIS_footer_year & " Total Grant: $" & Trim(total_grant) & "; Cash Portion: $" & Trim(cash_amt) & "; Food Portion: $" & Trim(food_amt) & "; Housing Grant: $" & Trim(housing_grant) & "; "
			Else 															'If ELIG results are from a different day
				Do 															'Goes to STAT SUMM to find if there is an inhibiting error
					CALL Navigate_to_MAXIS_screen ("STAT", "SUMM")
					EMReadScreen nav_check, 4, 2, 46
				Loop until nav_check = "SUMM"
				summ_row = 2
				Do 															'Reads each line on SUMM looking for CASH inhibited
					EMReadScreen edit_msg, 23, summ_row, 20
					If edit_msg = "CASH HAS BEEN INHIBITED" Then
						inhibiting_error = TRUE
						Exit do
					End If
					If trim(edit_msg) = "" Then 							'Goes to the next page of edits if needed
						EMReadScreen next_page, 7, summ_row, 71
						If next_page = "MORE: +" Then
							PF8
							summ_row = 1
						End If
					End iF
					summ_row = summ_row + 1
				Loop until summ_row = 23
				If inhibiting_error = TRUE then 							'Adds the inhibiting error information to the display
					MFIP_results = MFIP_results & MAXIS_footer_month & "/" & MAXIS_footer_year & " has an Inhibiting EDIT in STAT - resolve and rerun to generate results."
					inhibiting_error = FALSE								'Resets this value because each month needs to be assessed individually
				End If
			End IF
		Else 																'If no ELIG results
			Do																'Goes to STAT SUMM to find if there is an inhibiting error
				CALL Navigate_to_MAXIS_screen ("STAT", "SUMM")
				EMReadScreen nav_check, 4, 2, 46
			Loop until nav_check = "SUMM"
			summ_row = 2
			Do
				EMReadScreen edit_msg, 23, summ_row, 20
				If edit_msg = "CASH HAS BEEN INHIBITED" Then
					inhibiting_error = TRUE
					Exit do
				End If
				If trim(edit_msg) = "" Then
					EMReadScreen next_page, 7, summ_row, 71
					If next_page = "MORE: +" Then
						PF8
						summ_row = 1
					End If
				End iF
				summ_row = summ_row + 1
			Loop until summ_row = 23
			If inhibiting_error = TRUE then
				MFIP_results = MFIP_results & MAXIS_footer_month & "/" & MAXIS_footer_year & " has an Inhibiting EDIT in STAT - resolve and rerun to generate results."
				inhibiting_error = FALSE
			End If
		End If
	Next

	MAXIS_footer_month = right("00" & datepart("m", months_of_mfip_array(0)), 2)
	MAXIS_footer_year = right(datepart("yyyy", months_of_mfip_array(0)), 2)
	Back_to_SELF
	Call Navigate_to_MAXIS_screen ("ELIG", "MFIP")
	EMReadScreen elig_check, 4, 3, 47
	If elig_check = "MFPR" Then
		EMWriteScreen "MFSM", 20, 71
		transmit
	End If
End Function

FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date
	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

FUNCTION get_MFIP_case_info(ref_number, client_name)			'Function created to get basic information about a case
	Do 															'Starts at MEMB
		Call Navigate_to_MAXIS_screen("STAT", "MEMB")
		EMReadScreen nav_check, 4, 2, 48
	Loop until nav_check = "MEMB"

	If ref_number = "" Then EMReadScreen ref_number, 2, 4, 33	'Defaults to M01 if not defined
	ref_number = right("00" & ref_number, 2)					'Ref number must be 2 digits
	EMWriteScreen ref_number, 20, 76
	transmit
	EMReadScreen memb_on_case, 7, 8, 22							'Checks to make sure the Reference Number is used on this case
	If memb_on_case = "Arrival" Then
		PF3
		PF10													'If a wrong ref number was entered and this person does not exist on the case, the script will use M01
		MsgBox "HH Member " & ref_number & " does not exist on this case. The script will default to HH Member 01. Please check the reference number of the caregiver listed on the Status Update."
		ref_number = "01"
		EMWriteScreen ref_number, 20, 76
		transmit
	End If
	EMReadScreen first_name, 12, 6, 63							'Gets Client name and puts it together
	EMReadScreen last_name, 25, 6, 30

	first_name = Replace(first_name, "_", "")
	last_name = Replace(last_name, "_", "")
	client_name = first_name & " " & last_name & ""
	ref_number = ref_number & ""

	Call Navigate_to_MAXIS_screen ("STAT", "PROG")				'Goes to PROG to check if this is an MFIP case
	prog_row = 6
	Do 															'Cash has 2 lines on PROG and both should be checked
		EMReadScreen cash_prog, 2, prog_row, 67
		If cash_prog = "MF" Then 								'MFIP cases active or pending are allowed'
			EMReadScreen prog_status, 4, prog_row, 74
			If prog_status = "ACTV" or prog_status = "REIN" or prog_status = "PEND" Then
				Exit Do
			Else
				end_message = "This script is only for MFIP cases." & vbnewline & "MFIP is not Active, Pending or in REIN. The script will now end."
				script_end_procedure (end_message)
			End If
		ElseIf cash_prog = "  " Then 							'Sometimes the cash program is not defined in PROG while pending
			EMReadScreen prog_status, 4, prog_row, 74			'Will assume workers know this is a cash case
			If prog_status = "PEND" Then
				Exit Do
			Else
				end_message = "This script is only for MFIP cases." & vbnewline & "MFIP is not Active, Pending or in REIN. The script will now end."
				script_end_procedure (end_message)
			End If
		Else
			If prog_row = 7 Then
				end_message = "This script is only for MFIP cases." & vbnewline & "MFIP is not Active, Pending or in REIN. The script will now end."
				script_end_procedure (end_message)
			End If
		End If
		prog_row = prog_row + 1
	Loop Until prog_row = 8

	Call Navigate_to_MAXIS_screen ("STAT", "TIME")				'Needs to determine if this is an extension case or not
	EMWriteScreen ref_number, 20, 76
	transmit
	EMReadScreen MF_counted_mo, 2, 17, 69						'Need to remove the banked months from counted months for to asses true extension potential'
	EMReadScreen MF_banked_mo, 3, 19, 16
	MF_counted_mo = abs(MF_counted_mo)
	MF_banked_mo = abs(Trim(MF_banked_mo))
	tanf_ext = MF_counted_mo - MF_banked_mo
	If tanf_ext >= 60 Then
		end_message = "This script is for use on PRE-60 MFIP cases." & vbnewline & "This case has:" & vbnewline & MF_counted_mo & " counted TANF months." & vbnewline & MF_banked_mo & " banked TANF months." & vbnewline & "This case is not considered Pre-60 MFIP and has different process from FSS Coding. The script will now end."
		script_end_procedure(end_message)
	End IF
END FUNCTION

FUNCTION update_disa(ref_number, disa_start_date, disa_end_date, disa_status, disa_verif)
Do																'Function to write information to DISA
    Call Navigate_to_MAXIS_screen ("STAT", "DISA")				'Goes to DISA for the correct person
    EMReadScreen nav_check, 4, 2, 45
Loop until nav_check = "DISA"
EMWriteScreen ref_number, 20, 76
transmit
start_month = right("00" & DatePart("m", disa_start_date), 2)	'Isolates the start month, day, and year as these are seperate fields on DISA
start_day = right("00" & DatePart("d", disa_start_date), 2)
start_year = DatePart("yyyy", disa_start_date)

If IsDate(disa_end_date) = FALSE Then disa_end_date = DateAdd("m", 6, disa_start_date)
end_month = right("00" & DatePart("m", disa_end_date), 2)		'Isolates the end month, day, and year as these are seperate fields on DISA
end_day = right("00" & DatePart("d", disa_end_date), 2)
end_year = DatePart("yyyy", disa_end_date)
EMReadScreen disa_exist, 4, 6, 53								'Checking for dates already entered - looks at the DISA start year field
If disa_exist <> "____" Then 									'Anything listed here would indicate DISA is already loaded
    EMReadScreen listed_end_month, 2, 6, 69
    EMReadScreen listed_end_day, 2, 6, 72
    EMReadScreen listed_end_year, 4, 6, 75
    If listed_end_year = "____" Then disa_info = "It appears there is an open ended DISA for this person." 	'If no end date
    listed_end_date = listed_end_month & "/" & listed_end_day & "/" & listed_end_year
    listed_end_date = cDate(listed_end_date)
    If listed_end_date > date Then disa_info = "It appears there is DISA with a future end date for this person."
    If listed_end_date <= date Then disa_info = "It appears there is a DISA for this person that has already ended."	'WIll ask the user if the script should overwrite the current listed DISA dates
    change_disa_message = MsgBox(disa_info & vbNewLine & "Do you want the script to replace the dates on the panel with these?" & vbNewLine & vbNewLine & "Disability & Certification Begin: " & start_month & "/" & start_day & "/" & start_year & vbNewLine & "Disability & Certification End: " & end_month & "/" & end_day & "/" & end_year, vbYesNo + vbQuestion, "Update DISA?")
    If change_disa_message = VBNo Then panels_reviewed = panels_reviewed & "DISA for Memb " & ref_number & " & " ''
End If
If disa_exist = "____" or change_disa_message = VBYes Then		'If the panel is to be updated
    EMReadScreen numb_of_panels, 1, 2, 78						'Reading if it needs to create a new panel or just pF9
    IF numb_of_panels = "0" Then
        EMWriteScreen "NN", 20, 79
        transmit
    Else
        PF9
    End IF
    'Writing the Disability Begin Date'
    EMWriteScreen start_month, 6, 47
    EMWriteScreen start_day, 6, 50
    EMWriteScreen start_year, 6, 53
    'Writing the Certification Begin Date'
    EMWriteScreen start_month, 7, 47
    EMWriteScreen start_day, 7, 50
    EMWriteScreen start_year, 7, 53
    'Writing the Disability End Date'
    EMWriteScreen end_month, 6, 69
    EMWriteScreen end_day, 6, 72
    EMWriteScreen end_year, 6, 75
    'Writing the Certification End Date'
    EMWriteScreen end_month, 7, 69
    EMWriteScreen end_day, 7, 72
    EMWriteScreen end_year, 7, 75
    'Writing the verif code'
    EMWriteScreen disa_status, 11, 59
    EMWriteScreen disa_verif, 11, 69
    transmit
End If
END FUNCTION

FUNCTION Write_TIKL_if_needed (tikl_test, tikl_date, TIKL_message, TIKL_fail_masg)
	If tikl_test = TRUE Then
		If IsDate(tikl_date) = TRUE Then
			Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
			tikl_set_month = right("00" & DatePart("m", tikl_date), 2)
			tikl_set_day   = right("00" & DatePart("d", tikl_date), 2)
			tikl_set_year  = right(DatePart("yyyy", tikl_date), 2)

			EMWriteScreen tikl_set_month, 5, 18
			EMWriteScreen tikl_set_day,   5, 21
			EMWriteScreen tikl_set_year,  5, 24
			transmit

			Call Write_variable_in_TIKL (TIKL_message)
			transmit
			EMReadScreen TIKL_verified, 4, 24, 2
			IF TIKL_verified = "    " Then
				tikl_test = TRUE
			ELSE
				tikl_test = FALSE
				MsgBox TIKL_fail_masg
			End If
			PF3
			Back_to_SELF
		Else
			MsgBox TIKL_fail_masg
		End If
	End If
END FUNCTION
'===============================================================================================================================

EMConnect ""

IF worker_county_code = "x162" Then Ramsey_County_case = TRUE
Call MAXIS_case_number_finder(MAXIS_case_number)		'Looks for a case number

If MAXIS_case_number <> "" Then							'If one is found, will get case information using M01 as default
	CALL get_MFIP_case_info (ref_number, client_name)
End IF

If client_name = "" then client_name = "Enter Ref Numb and press 'Reload Client Name'"	'Loads instructions into the edit box

'Running the first dialog
Do
	err_msg = ""
	Dialog fss_status_dialog
	Cancel_confirmation
	If ButtonPressed = get_client_name_button Then CALL get_MFIP_case_info (ref_number, client_name)
	If universal_partipant_checkbox = unchecked AND new_imig_checkbox = unchecked AND age_sixty_checkbox = unchecked AND preg_checkbox = unchecked AND ill_incap_checkbox = unchecked AND care_of_ill_Incap_checkbox = unchecked AND child_under_one_checkbox = unchecked AND fam_violence_checkbox = unchecked AND Special_medical_checkbox = unchecked AND iq_test_checkbox = unchecked AND learning_disabled_checkbox = unchecked AND mentally_ill_checkbox = unchecked AND ssi_pending_checkbox = unchecked AND unemployable_checkbox = unchecked AND dev_delayed_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "You must select a code to update."
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a case number."
	If ref_number = "" Then err_msg = err_msg & vbNewLine & "Please enter the reference number of the person the SU is for."
	If err_msg <> "" AND ButtonPressed <> get_client_name_button Then MsgBox "Please resolve to continue." & vbNewLine & err_msg
Loop until err_msg = "" AND ButtonPressed = OK

Call check_for_maxis(False)

ref_number = right("00" & ref_number, 2)			'Makes reference number 2 digit
Call Navigate_to_MAXIS_screen("STAT", "EMPS")		'Goes to EMPS for the designated person
EMWriteScreen ref_number, 20, 76
transmit
EMReadScreen current_emps_status, 38, 15, 40		'Find the emps status
current_emps_status = trim(current_emps_status)
CALL Navigate_to_MAXIS_screen ("STAT", "MEMI")
EMReadScreen current_fvw, 2, 17, 78					'Reads the code if there is a family violence waiver currently coded

new_fvw = TRUE 										'setting the default for these variables
new_fss = FALSE

If current_emps_status = "20 (UP) Universal Participation" Then new_fss = TRUE	'If currently a universal partcipant, then the FSS coding is New
If current_fvw = "02" Then new_fvw = FALSE 			'If waiver code is 02 on MEMI then the FVW is a renewal

If child_under_one_checkbox = checked then 			'Looking for a baby on the case if child under 1 is checked
	baby_on_case = FALSE							'Defaults to false
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "PNLP")
		EMReadScreen nav_check, 4, 2, 53
	Loop until nav_check = "PNLP"
	maxis_row = 3
	Do
		EMReadScreen panel_name, 4, maxis_row, 5	'Reads the name of each panel listed on PNLP
		If panel_name = "MEMB" Then 				'Looking for MEMB
			EMReadScreen client_age, 2, maxis_row, 71		'Reads the age on the MEMB line
			If client_age = " 0" Then
				baby_on_case = TRUE	'If a age is listed as 0 then a baby is on the case'
				EMReadScreen Baby_ref_numb, 2, 10, 10
			End If
		End If
		If panel_name = "MEMI" Then Exit Do			'Once it gets to a panel named MEMI, there are no additional MEMB panels
		maxis_row = maxis_row + 1					'Go to next row
		If maxis_row = 20 Then 						'If it gets to row 20 it needs to go to the next page
			transmit
			maxis_row = 3
		End If
	Loop until panel_name = "REVW"
	If baby_on_case = FALSE Then 		'If there is no baby on the case the script will not update to a child under 12 months exemption - this notifies the worker and unchecks the selector
		no_baby_message = MsgBox("You are reporting an FSS status with a child under 12 months but there is no child under 1 listed in the household. You must add the baby fisrt." & vbNewLine & "Press Cancel to stop the script." & vbNewLine & "Press OK to continue the script if you have selected other FSS reasons.", vbOKCancel + VBAlert, "Review Child Under 12 Months Selection")
		If no_baby_message = VBCancel then cancel_confirmation
		child_under_one_checkbox = unchecked
	Else
		Call Navigate_to_MAXIS_screen ("STAT", "MEMB")
		EMWriteScreen Baby_ref_numb, 20, 76
		transmit
		EMReadScreen Baby_DOB, 10, 8, 42
		Baby_DOB = replace(Baby_DOB, " ", "/")
		Baby_is_One = DateAdd("yyyy", 1, Baby_DOB)
		Exemption_Unaavailable = DateAdd("m", 1, Baby_is_One)
		Exemption_End_Month = right("00" & DatePart("m", Exemption_Unaavailable), 2)
		Exemption_End_Year = DatePart("yyyy", Exemption_Unaavailable)
	End If
End If

If child_under_one_checkbox = checked Then 				'If child under 1 is requested, go to EMPS to figure out which months have already been used
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen "X", 12, 39							'Open the list of exemption months already taken
	transmit
	emps_row = 7										'Setting the first row and col
	emps_col = 22
	Do
		EMReadScreen month_used, 2, emps_row, emps_col	'reading the first field
		If month_used = "__" Then Exit Do				'if the month was listed as blank, there are no more months listed
		EMReadScreen year_used, 4, emps_row, emps_col + 5		'reads the year associated with the month listed
		emps_exemption_month_used = emps_exemption_month_used & "~" & month_used & "/" & year_used	'adds the month and year to a string seperated by ~
		emps_col = emps_col + 11						'moves to the next month listed spot
		If emps_col = 66 Then 							'Once it has gone through all the fields on this row, it goes to the next row and starts over at the beginning of the columns.
			emps_col = 22
			emps_row = emps_row + 1
		End If
	Loop Until emps_row = 10							'There are only 3 rows of data
	If emps_exemption_month_used <> "" Then
		emps_exemption_month_used = right(emps_exemption_month_used, len(emps_exemption_month_used)-1)	'lops off the extra ~ at the beginning
		used_expemption_months_array = split(emps_exemption_month_used, "~")							'creates an array for the counting
		months_for_use = Join(used_expemption_months_array, ", ")										'creates a string of months used for case noting
		number_of_months_available = 12 - (ubound(used_expemption_months_array) + 1) & ""				'uses the ubound of the array to determine how many months are left to be used
	Else
		months_for_use = "NONE"
		number_of_months_available = 12
	End If
End If

'This is required to set the size of the next dialog based on the checkboxes on the previous dialog
months_to_fill = "Enter the date of request and click 'Calculate' to fill this field."				'instructions in the edit box
detail_dialog_length = 45
If ill_incap_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If care_of_ill_Incap_checkbox = checked Then detail_dialog_length = detail_dialog_length + 60
If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If fam_violence_checkbox = checked Then
	detail_dialog_length = detail_dialog_length + 40
	fvw_only = TRUE
End IF
If ssi_pending_checkbox = checked Then detail_dialog_length = detail_dialog_length + 40
If child_under_one_checkbox = checked Then detail_dialog_length = detail_dialog_length + 60
If new_imig_checkbox = checked Then detail_dialog_length = detail_dialog_length + 35
If Special_medical_checkbox = checked Then detail_dialog_length = detail_dialog_length + 60
y_pos_counter = 25

'This is the second Dialog - which is defined here because it is dynamic.
BeginDialog fss_code_detail, 0, 0, 440, detail_dialog_length, "Update FSS Information from the Status Update"
  Text 5, 10, 40, 10, "Date of SU"
  EditBox 50, 5, 50, 15, SU_date
  Text 120, 10, 40, 10, "ES Agency"
  EditBox 165, 5, 65, 15, es_agency
  Text 240, 10, 40, 10, "ES Worker"
  EditBox 285, 5, 110, 15, es_worker

  If ill_incap_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 35, "Client Illness/Incapacity"
	  Text 15, y_pos_counter + 20, 40, 10, "Start Date"
	  EditBox 75, y_pos_counter + 15, 50, 15, ill_incap_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, ill_incap_end_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", ill_incap_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", ill_incap_docs_with_fas

	  y_pos_counter = y_pos_counter + 40
  End If
  If care_of_ill_Incap_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 55, "Needed in Home to care for Family Member"
	  Text 15, y_pos_counter + 20, 95, 10, "Person in HH requiring care"
	  EditBox 115, y_pos_counter + 15, 25, 15, disa_HH_memb
	  Text 15, y_pos_counter + 40, 55, 10, "DISA Start Date"
	  EditBox 75, y_pos_counter + 35, 50, 15, rel_care_start_date
	  Text 135, y_pos_counter + 40, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 35, 50, 15, rel_care_end_date
	  Text 260, y_pos_counter + 40, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 40, 25, 10, "ES", rel_care_docs_with_es
	  CheckBox 370, y_pos_counter + 40, 50, 10, "Financial", rel_care_docs_with_fas

	  y_pos_counter = y_pos_counter + 60
  End If
  If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 35, "Unemployable"
	  Text 15, y_pos_counter + 20, 55, 10, "Start Date on SU"
	  EditBox 75, y_pos_counter + 15, 50, 15, unemployable_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, unemployable_end_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", unemployable_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", unemployable_docs_with_fas

	  y_pos_counter = y_pos_counter + 40
  End If
  If fam_violence_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 35, "Family Violence Waiver"
	  Text 15, y_pos_counter + 20, 55, 10, "Start Date "
	  EditBox 75, y_pos_counter + 15, 50, 15, fvw_start_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, fvw_end_date

	  y_pos_counter = y_pos_counter + 40
  End If
  If ssi_pending_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 35, "SSI/RSDI Pending"
	  Text 15, y_pos_counter + 20, 55, 10, "Application Date"
	  EditBox 75, y_pos_counter + 15, 50, 15, ssa_app_date
	  Text 135, y_pos_counter + 20, 35, 10, "End Date"
	  EditBox 180, y_pos_counter + 15, 50, 15, ssa_end_date
	  Text 260, y_pos_counter + 20, 70, 10, "Documentation with:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", ssa_app_docs_with_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", ssa_app_docs_with_fas

	  y_pos_counter = y_pos_counter + 40
  End If
  If child_under_one_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 55, "Child Under 12 Months"
	  Text 15, y_pos_counter + 20, 55, 10, "Request Date"
	  EditBox 75, y_pos_counter + 15, 50, 15, child_under_1_request_date
	  Text 275, y_pos_counter + 20, 85, 10, "Request made to:"
	  CheckBox 335, y_pos_counter + 20, 25, 10, "ES", child_under_1_at_es
	  CheckBox 370, y_pos_counter + 20, 50, 10, "Financial", child_under_1_at_fas
	  Text 130, y_pos_counter + 10, 65, 10, "Months used:"
	  Text 175, y_pos_counter + 10, 100, 30, months_for_use
	  Text 15, y_pos_counter + 40, 75, 10, "Months of exemption"
	  EditBox 90, y_pos_counter + 35, 280, 15, months_to_fill
	  ButtonGroup ButtonPressed
	    PushButton 380, y_pos_counter + 40, 35, 10, "Calculate", child_under_1_months_calculate

	  y_pos_counter = y_pos_counter + 60
  End If
  If new_imig_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 430, 50, "Newly Arrived Immigrant"
	  Text 15, y_pos_counter + 15, 110, 10, "Spoken Language (SPL) from SU"
	  EditBox 130, y_pos_counter + 10, 25, 15, spl_listed
	  CheckBox 170, y_pos_counter + 15, 260, 10, "Check here to confirm that the SU indicates clt is enrolled in ELL/ESL classes", ell_confirm_checkbox
	  Text 70, y_pos_counter + 35, 35, 10, "End Date"
	  EditBox 105, y_pos_counter + 30, 50, 15, new_imig_end_date
	  Text 170, y_pos_counter + 35, 205, 10, "If ledt blank a TIKL to review will be set for 6 months from now."

	  y_pos_counter = y_pos_counter + 55
  End If
  If Special_medical_checkbox = checked Then
	  GroupBox 5, y_pos_counter, 440, 45, "Special Medical Criteria"
	  Text 20, y_pos_counter + 15, 100, 10, "Person in HH meeting Criteria"
	  EditBox 125, y_pos_counter + 10, 20, 15, smc_hh_memb
	  Text 160, y_pos_counter + 15, 70, 10, "Date of Diagnosis"
	  EditBox 225, y_pos_counter + 10, 50, 15, smc_diagnosis_date
	  Text 310, y_pos_counter + 15, 30, 10, "End Date"
	  EditBox 345, y_pos_counter + 10, 50, 15, smc_end_date
	  Text 65, y_pos_counter + 35, 60, 10, "Medical Criteria"
	  DropListBox 125, y_pos_counter + 30, 125, 40, "Select One ..."+chr(9)+"1 - Home-Health/Waiver Services"+chr(9)+"2 - Child who meets SED Criteria"+chr(9)+"3 - other Adult who meets SPMI", medical_criteria
	  Text 275, y_pos_counter + 35, 70, 10, "Documentation with:"
	  CheckBox 350, y_pos_counter + 35, 25, 10, "ES", smp_docs_with_es
	  CheckBox 385, y_pos_counter + 35, 50, 10, "Financial", smc_docs_with_fas

	  y_pos_counter = y_pos_counter + 60
  End If

  Text 15, y_pos_counter, 85, 15, "Caregiver SU received for:"
  Text 100, y_pos_counter, 150, 15, client_name
  ButtonGroup ButtonPressed
	OkButton 330, y_pos_counter, 50, 15
	CancelButton 385, y_pos_counter, 50, 15
EndDialog

'Calls the second dialog
Do
	Do
		err_msg = ""
		dialog fss_code_detail
		cancel_confirmation
		If ButtonPressed = child_under_1_months_calculate Then
			If IsDate(child_under_1_request_date) = TRUE Then 	'This creates a list of the future months to be coded as exempt on EMPS.
				For add_month = 1 to number_of_months_available		'using the count determined in the EMPS
					this_month = DatePart("m", DateAdd ("m", add_month, child_under_1_request_date))	'first month is the month after the exemption is requested, then adding all the others after'
					If len(this_month) = 1 Then this_month = "0" & this_month		'making 2 digit
					this_year = DatePart("yyyy", DateAdd("m", add_month, child_under_1_request_date))	'creating a year
					If trim(this_month) = trim(Exemption_End_Month) AND trim(this_year) = trim(Exemption_End_Year) Then Exit For
					new_exemption_months = new_exemption_months & "~" & this_month & "/" & this_year	'list of all of these months
				Next
				If new_exemption_months <> "" Then
					new_exemption_months = right(new_exemption_months, len(new_exemption_months) - 1)		'taking off the extra ~
					new_exemption_months_array = split(new_exemption_months, "~")							'creating an array of the months to code for future exempt months
					months_to_fill = Join(new_exemption_months_array, ", ")									'list for the edit box
					Impose_Exemption = TRUE
				Else
					months_to_fill = "None available."
					MsgBox "It appears the baby on this case will turn one before the Child Under One Exemption can be put into place. If you contine the script with this date as the request date, this exemption will not be coded. Otherwise review the request date."
					Impose_Exemption = FALSE
				End If
			Else
				MsgBox "You must enter a valid date to calculate the which months will have an exemption."
			End If
		End If
		If es_worker = "" Then err_msg = err_msg & vbNewLine & "** You must enter the name of the ES worker that completed the SU."
		If es_agency = "" Then err_msg = err_msg & vbNewLine & "** You must enter the ES Agency that provided the SU."
		If IsDate(SU_date) = FALSE Then err_msg = err_msg & vbNewLine & "** Enter the date of the Status Update."
		If ill_incap_checkbox = checked Then
			If IsDate(ill_incap_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start of client Ill/Incap. If one was not provided on the SU, an new SU is required."
			If ill_incap_docs_with_es = unchecked AND ill_incap_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of client's ill/incap are held in ES file or Financial File."
		End If
		If care_of_ill_Incap_checkbox = checked Then
			If IsNumeric(disa_HH_memb) = False Then err_msg = err_msg & vbNewLine & "- List the reference number of the household member the client is needed in the home to care for. The person must be listed on the case, if the person has not yet been added to the case, cancel the script and do that first."
			If IsDate(rel_care_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start need to be at home. If one was not provided on the SU, an new SU is required."
			If rel_care_docs_with_es = unchecked AND rel_care_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of need to be at home for care of a family member is held in ES file or Financial File."
		End If
		If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then
			If IsDate(unemployable_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the start of client determined to be unemployable. If one was not provided on the SU, an new SU is required."
			If IsDate(unemployable_start_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date for the end of client determined to be unemployable. If one was not provided on the SU, an new SU is required."
			If unemployable_docs_with_es = unchecked AND unemployable_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of client's unemployability is held in ES file or Financial File."
		End If
		If fam_violence_checkbox = checked Then
			If IsDate(fvw_start_date) = False Then err_msg = err_msg & vbNewLine & "- Start date of Family Violence Waiver must be listed. If one was not provided on the SU, an new SU is required."
			If IsDate(fvw_end_date) = False Then err_msg = err_msg & vbNewLine & "- End date of Family Violence Waiver must be listed. If one was not provided on the SU, an new SU is required."
		End If
		If ssi_pending_checkbox = checked Then
			If IsDate(ssa_app_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a valid date of applicaiton for SSI/RSDI."
			If ssa_app_docs_with_es = unchecked AND ssa_app_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of client's SSI/RSDI Application is held in ES file or Financial File"
		End If
		If child_under_one_checkbox = checked Then
			If IsDate(child_under_1_request_date) = False Then err_msg = err_msg & vbNewLine &"- You must enter a the date the Child Under 12 Months Exemption was requested."
			If child_under_1_at_es = unchecked AND child_under_1_at_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if the request for Child Under 12 Months Exemption was requested to ES or Financial."
		End If
		If new_imig_checkbox = checked Then
			If ell_confirm_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "- The SU must confirm that clt is enrolled in ELL Classes, if it does not a new SU is required."
			spl_listed = abs(spl_listed)
			If spl_listed >= 6 Then err_msg = err_msg & vbNewLine & "- Spoken Language (SPL) must be less than 6 to qualify for this FSS Coding. Connect with ES worker to clarify."
		End If
		If Special_medical_checkbox = checked Then
			If IsNumeric(smc_hh_memb) = False Then err_msg = err_msg& vbNewLine & "- List the reference number of the household member who qualifies for Special Medical Criteria. The person must be listed on the case, if the person has not yet been added to the case, cancel the script and do that first."
			If IsDate(smc_diagnosis_date) = False Then MsgBox "No Diagnosis Date was listed, it is not required, but TANF Banked Months cannot be determined without it."
			If smp_docs_with_es = unchecked AND smc_docs_with_fas = unchecked Then err_msg = err_msg & vbNewLine & "- Please indicate if verification of need to be at home for care of a family member is held in ES file or Financial File."
			If medical_criteria = "Select One ..." Then err_msg = err_msg & "- Select a Medical Criteria from what is indicated on the SU."
		End If
		If err_msg <> "" AND ButtonPressed <> child_under_1_months_calculate Then MsgBox "You must resolve to continue:" & vbNewLine & vbNewLine & err_msg
	Loop until err_msg = "" AND ButtonPressed = OK
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

If Impose_Exemption = FALSE Then child_under_one_checkbox = unchecked

Back_to_SELF
'The footer month/year defaults to the status update date for most of the categories
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If ill_incap_checkbox = checked Then 		'FSS CATEGORY - CLIENT ILL OR INCAPACITATED
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 							'This variable determines which case notes will happen later
	ill_incap_tikl = TRUE
	ill_incap_tikl_date = ill_incap_end_date
	fss_category_list = fss_category_list & "; Ill/Incap >60 Days"		'Creates a list of all the categories for case notes
	If ill_incap_end_date = "" Then ill_incap_end_date = DateAdd("m", 6, ill_incap_start_date)		'Defaults the end date of ill/incap to 6 months from the start date
	CALL update_disa(ref_number, ill_incap_start_date, ill_incap_end_date, "09", "6")		'Calls the function to update the DISA panel
	panels_updated = panels_updated & "DISA for Memb " & ref_number & " & "		'Creates a list of panels updated for case notes
End If

If care_of_ill_Incap_checkbox = checked Then 		'FSS CATEGORY - CLIENT IS REQUIRED IN HOME TO CARE FOR ILL OR INCAPACITATED HOUSEHOLD MEMBER
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 							'This variable determines which case notes will happen later
	care_of_ill_incap_tikl = TRUE
	care_of_ill_incap_tikl_date = rel_care_end_date
	fss_category_list = fss_category_list & "; Care of Ill/Incap Family Member"		'Creates a list of all the categories for case notes
	If rel_care_end_date = "" Then rel_care_end_date = DateAdd("m", 6, rel_care_start_date)		'Defaults the end date of relative care need to 6 months from the start date if none is defined
	CALL update_disa(disa_HH_memb, rel_care_start_date, rel_care_end_date, "09", "6")			'Calls the function to update the DISA panel
	panels_updated = panels_updated & "DISA for Memb " & disa_HH_memb & " & "			'Creates a list of panels updated for case notes
End If

'FSS CATEGORY - CLIENT MEETS ONE OF THE UNEMPLOYABLE CATEGORIES
If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 							'This variable determines which case notes will happen later
	hard_to_employ_tikl = TRUE
	hard_to_employ_tikl_date = unemployable_end_date
	fss_category_list = fss_category_list & "; Hard to Employ"		'Creates a list of all the categories for case notes
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")				'Navigates to EMPS and updates based on which unemployable category was selected
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	PF9
	If unemployable_checkbox = checked Then
		EMWriteScreen "UN", 11, 76
		fss_category_list = fss_category_list & " - Unemployable"	'Specifics of the unemployable category are added to the list of FSS Category for case note
	End If
	If dev_delayed_checkbox = checked Then
		EMWriteScreen "DD", 11, 76
		fss_category_list = fss_category_list & " - Developmentally Delayed"	'Specifics of the unemployable category are added to the list of FSS Category for case note
		update_disa_for_UN = TRUE
	End If
	If mentally_ill_checkbox = checked Then
		EMWriteScreen "MI", 11, 76
		fss_category_list = fss_category_list & " - Mentally Ill"			'Specifics of the unemployable category are added to the list of FSS Category for case note
		update_disa_for_UN = TRUE
	End If
	If learning_disabled_checkbox = checked Then
		EMWriteScreen "LD", 11, 76
		fss_category_list = fss_category_list & " - Learning Disabled"			'Specifics of the unemployable category are added to the list of FSS Category for case note
		update_disa_for_UN = TRUE
	End If
	IF iq_test_checkbox = checked Then
		EMWriteScreen "IQ", 11, 76
		fss_category_list = fss_category_list & " - IQ Tested < 80"			'Specifics of the unemployable category are added to the list of FSS Category for case note
		update_disa_for_UN = TRUE
	End If
	transmit
	panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes
	If update_disa_for_UN = TRUE Then
		CALL update_disa(ref_number, unemployable_start_date, unemployable_end_date, "09", "6")
		panels_updated = panels_updated & "DISA for Memb " & ref_number & " & "		'Creates a list of panels updated for case notes
	End IF
End If

If fam_violence_checkbox = checked Then 				'FSS CATEGORY - HOUSEHOLD HAS A FAMILY VIOLENCE WAIVER
	STATS_counter = STATS_counter + 1
	fss_category_list = fss_category_list & "; Family Violence Waiver"		'Creates a list of all the categories for case notes
	fvw_tikl = TRUE
	fvw_tikl_date = fvw_end_date
	MAXIS_footer_month = right("00" & DatePart("m", fvw_start_date), 2)			'Case needs to be updated in the footer month that the waiver starts
	MAXIS_footer_year = right(DatePart("yyyy", fvw_start_date), 2)
	Back_to_SELF
	Do
		Do															'Goes to MEMI for the person to code the family violence waiver
			Call Navigate_to_MAXIS_screen ("STAT", "MEMI")
			EMReadScreen nav_check, 4, 2, 50
		Loop until nav_check = "MEMI"
		EMWriteScreen ref_number, 20, 76
		transmit
		PF9
		EMWriteScreen "02", 17, 78								'Adding the code and the start month/year
		EMWriteScreen MAXIS_footer_month, 18, 49
		EMWriteScreen MAXIS_footer_year, 18, 55
		transmit
		transmit
		PF3														'The MEMI panel needs to be checked in every subsequent year to make sure the data did not expire
		EMWriteScreen "Y", 16, 54
		transmit
		EMReadScreen end_wrap, 24, 24, 2
	Loop until end_wrap = "CONTINUATION NOT ALLOWED"			'This is what is on the STAT WRAP screen in CM plus 1

	panels_updated = panels_updated & "MEMI for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes

	next_month = DateAdd("m", 1, date)							'setting up the date variables needed to loop through all of the fields on time
	next_mo = right("00" & DatePart("m", next_month) , 2)
	next_yr = right(DatePart("yyyy", next_month), 2)
	next_MAXIS_month = next_mo & "/" & next_yr
	Do															'Go to STAT TIME
		Call Navigate_to_MAXIS_screen ("STAT", "TIME")
		EMReadScreen nav_check, 4, 2, 46
	Loop until nav_check = "TIME"
	EMWriteScreen ref_number, 20, 76							'For the person the SU is for
	transmit
	Do
		If MAXIS_footer_month = "01" Then fvw_month_col = 15	'Defining where all the fields are on the TIME panel.
		If MAXIS_footer_month = "02" Then fvw_month_col = 20
		If MAXIS_footer_month = "03" Then fvw_month_col = 25
		If MAXIS_footer_month = "04" Then fvw_month_col = 30
		If MAXIS_footer_month = "05" Then fvw_month_col = 35
		If MAXIS_footer_month = "06" Then fvw_month_col = 40
		If MAXIS_footer_month = "07" Then fvw_month_col = 45
		If MAXIS_footer_month = "08" Then fvw_month_col = 50
		If MAXIS_footer_month = "09" Then fvw_month_col = 55
		If MAXIS_footer_month = "10" Then fvw_month_col = 60
		If MAXIS_footer_month = "11" Then fvw_month_col = 65
		If MAXIS_footer_month = "12" Then fvw_month_col = 70
		For row = 5 to 16										'Looks at the years listed on the left to find the year of the first month to check '
			EMReadScreen find_year, 2, row, 11
			If MAXIS_footer_year = find_year Then 				'If the year matches, it looks at the col defined above for the month
				fvw_month_row = row
				Exit For
			End If
		Next
		EMReadScreen is_counted, 2, fvw_month_row, fvw_month_col	'Reading the current code at this place in the TIME panel.
		If is_counted = "SS" OR is_counted = "SF" OR is_counted = "WS" OR is_counted = "WF" Then 		'If a counted month is coded
			PF9 																						'Edit mode
			EMWriteScreen "WD", fvw_month_row, fvw_month_col											'Write the FVW code
			counted_months_changed = counted_months_changed & " & " & MAXIS_footer_month & "/" & MAXIS_footer_year	'Creates a list of the months on TIME that were changed for case note
		End If
		Call month_change(1, MAXIS_footer_month, MAXIS_footer_year, month_ahead, month_yr_ahead)		'Goes to the next month for the loop
		MAXIS_footer_month = month_ahead																'Resets the footer month and year for the next loop
		MAXIS_footer_year = month_yr_ahead
	Loop until MAXIS_footer_month & "/" & MAXIS_footer_year = next_MAXIS_month		'compares the month that is being reviewed to the variable set above for CM + 1
	transmit
	panels_updated = panels_updated & "TIME for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes
	EMReadScreen tanf_used, 3, 17, 69
	EMReadScreen ext_tanf_used, 3, 19, 69
End If

tanf_used = trim(tanf_used)
ext_tanf_used = trim(ext_tanf_used)
If counted_months_changed <> "" Then counted_months_changed = right (counted_months_changed, len(counted_months_changed)-3)		'reformatting the lsit of months changed for case note

Back_to_SELF	'Resetting the footer month because FVW changed it
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If ssi_pending_checkbox = checked Then 					'FSS CATEGORY - SSI/RSDI ARE PENDING
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 									'This variable determines which case notes will happen later
	ssi_pending_tikl = TRUE
	fss_category_list = fss_category_list & "; SSI/RSDI Pending"		'Creates a list of all the categories for case notes
	MAXIS_footer_month = right("00" & DatePart("m", ssa_app_date), 2)	'Needs to be coded in the month that the app happened
	MAXIS_footer_year = right(DatePart("yyyy", ssa_app_date), 2)
	Back_to_SELF
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "PBEN")					'Go to PBEN for the correct client
		EMReadScreen nav_check, 4, 2, 49
	Loop until nav_check = "PBEN"
	ssa_app_month = right("00" & DatePart("m", ssa_app_date), 2)		'Setting up the parts of the date for MXIS fields
	ssa_app_day = right("00" & DatePart("d", ssa_app_date), 2)
	ssa_app_year = right(DatePart("yyyy", ssa_app_date), 2)
	pben_row = 8 														'Setting the starting point for reading the whole panel
	Do
		EMReadScreen pben_exist, 2, pben_row, 24
		If pben_exist = "__" Then 										'looks at the first line and if the type is blank, nothing is listed here
			EMReadScreen numb_of_panels, 1, 2, 78
			IF numb_of_panels = "0" Then 								'putting the panel in edit mode
				EMWriteScreen "NN", 20, 79
				transmit
			Else
				PF9
			End IF
			EMWriteScreen "01", pben_row, 24							'Adding the SSI app and RSDI app codes to pben as pending
			EMWriteScreen ssa_app_month, pben_row, 51
			EMWriteScreen ssa_app_day, pben_row, 54
			EMWriteScreen ssa_app_year, pben_row, 57
			EMWriteScreen "5", pben_row, 62
			EMWriteScreen "P", pben_row, 77

			EMWriteScreen "02", pben_row + 1, 24
			EMWriteScreen ssa_app_month, pben_row + 1, 51
			EMWriteScreen ssa_app_day, pben_row + 1, 54
			EMWriteScreen ssa_app_year, pben_row + 1, 57
			EMWriteScreen "5", pben_row + 1, 62
			EMWriteScreen "P", pben_row + 1, 77

			panels_updated = panels_updated & "PBEN for Memb " & ref_number & " & "
			Exit Do
		ElseIf pben_exist = "01" OR pben_exist = "02" Then 			'if SSI or RSDI are already listed as pending the script will ask if the worker wants to replace it
			EMReadScreen listed_app_month, 2, pben_row, 51
			EMReadScreen listed_app_day, 2, pben_row, 54
			EMReadScreen listed_app_year, 2, pben_row, 57
			IF listed_app_month = ssa_app_month AND listed_app_day = ssa_app_day AND listed_app_year = ssa_app_year Then 	'Asking the worker if PBEN is correct
				same_pben_date_msg = MsgBox ("It appears this SSI/RSDI Application information is already listed on PBEN." & vbNewLine & "Review the application information listed on PBEN." & vbNewLine & "Are the SSI and RSDI lines both listed correctly?", vbYesNo + vbQuestion, "PBEN data duplicated?")
				If same_pben_date_msg = vbYes then 	'answering yes says the PBEN is correct and so does not need to be updated'
					panels_reviewed = panels_reviewed & "PBEN for Memb " & ref_number & " - SSI/RSDI application already listed & "
					Exit Do
				ElseIf same_pben_date_msg = vbNo then
					EMReadScreen next_pben_exist, 2, pben_row + 1, 24
					IF next_pben_exist = "__" OR next_pben_exist = "01" OR next_pben_exist = "02" Then
						PF9																					'Adding the SSI app and RSDI app codes to pben as pending
						EMWriteScreen "01", pben_row, 24
						EMWriteScreen ssa_app_month, pben_row, 51
						EMWriteScreen ssa_app_day, pben_row, 54
						EMWriteScreen ssa_app_year, pben_row, 57
						EMWriteScreen "5", pben_row, 62
						EMWriteScreen "P", pben_row, 77

						EMWriteScreen "02", pben_row + 1, 24
						EMWriteScreen ssa_app_month, pben_row + 1, 51
						EMWriteScreen ssa_app_day, pben_row + 1, 54
						EMWriteScreen ssa_app_year, pben_row + 1, 57
						EMWriteScreen "5", pben_row + 1, 62
						EMWriteScreen "P", pben_row + 1, 77

						panels_updated = panels_updated & "PBEN for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes
						Exit Do
					Else
						panels_reviewed = panels_reviewed & "PBEN for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes
						MsgBox "PBEN could not be updated, and will need to be updated manually"
					End IF
				End If
			End IF
		Else
			pben_row = pben_row + 1																			'Or go to the next line
		End If
	Loop until pben_row = 12
	If pben_row = 12 Then replace_pben_message = MSGBox("It appears the PBEN Panel is full." & vbNewLine & vbNewLine & "The script can overwrite the first 2 lines with the pending SSI/RSDI application." & vbNewLine & vbNewLine & "If you agree to application information being entered on the first 2 lines, press 'Yes'", vbYesNo + vbAlert, "Update PBEN?")
	If replace_pben_message = vbYes Then 			'If PBEN is full of things other than SSI and RSDI and the worker wants to replace the first two lines, the script can overwrite
		PF9											'Rewriting the top two lines with SSI/RSDI pending
		EMWriteScreen "01", 8, 24
		EMWriteScreen ssa_app_month, 8, 51
		EMWriteScreen ssa_app_day, 8, 54
		EMWriteScreen ssa_app_year, 8, 57
		EMWriteScreen "5", 8, 62
		EMWriteScreen "P", 8, 77

		EMWriteScreen "02", 9, 24
		EMWriteScreen ssa_app_month, 9, 51
		EMWriteScreen ssa_app_day, 9, 54
		EMWriteScreen ssa_app_year, 9, 57
		EMWriteScreen "5", 9, 62
		EMWriteScreen "P", 9, 77

		panels_updated = panels_updated & "PBEN for Memb " & ref_number & " & "		'Creates a list of panels updated for case notes
	ElseIF replace_pben_message = vbNo Then
		panels_reviewed = panels_reviewed & "PBEN for Memb " & disa_HH_memb & " & "
	End If
	transmit
	If ssa_end_date = "" Then ssa_end_date = DateAdd("m", 6, ssa_app_date)
	CALL update_disa (ref_number, ssa_app_date, ssa_end_date, "06", "6")			'Updating DISA with the pending application dates
	panels_updated = panels_updated & "DISA for Memb " & ref_number & " & "				'Creates a list of panels updated for case notes
	ssi_pending_tikl_date = ssa_end_date
End If

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If child_under_one_checkbox = checked Then 						'FSS CATEGORY - CAREGIVER OF A CHILD UNDER 12 MONTHS
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 											'This variable determines which case notes will happen later
	child_under_one_tikl = TRUE
	last_month = left(new_exemption_months_array(ubound(new_exemption_months_array)), 2)
	last_year = right(new_exemption_months_array(ubound(new_exemption_months_array)), 2)
	child_under_one_tikl_date = last_month & "/01/" & last_year
	fss_category_list = fss_category_list & "; Care of Child < 12 Months"		'Creates a list of all the categories for case notes
	MAXIS_footer_month = left(new_exemption_months_array(0), 2)		'getting footer month by using the array of months to be exempt
	MAXIS_footer_year = right(new_exemption_months_array(0), 2)
	Do															'Go to EMPS
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen ref_number, 20, 76
	transmit
	PF9
	EMWriteScreen "Y", 12, 76
	EMWriteScreen "X", 12, 39
	transmit

	emps_row = 7												'setting the first location
	emps_col = 22
	Do
		EMReadScreen month_used, 2, emps_row, emps_col			'finding the first blank month to code
		If month_used = "__" Then Exit Do
		emps_col = emps_col + 11
		If emps_col = 66 Then
			emps_col = 22
			emps_row = emps_row + 1
		End If
	Loop Until emps_row = 10
	IF emps_row = 10 Then 										'if there are no blank months then error - cannot code an exemption
		MsgBox "It appears the client has used all of their Exempt Months. EMPS will need to be updated manually."
		PF3
		PF10
	Else
		For each exempt_month in new_exemption_months_array				'writing each of the months to be exempt in the array into the popup
			EMWriteScreen left(exempt_month, 2), emps_row, emps_col
			EMWriteScreen right(exempt_month, 4), emps_row, emps_col + 5
			emps_col = emps_col + 11
			If emps_col = 66 Then
				emps_col = 22
				emps_row = emps_row + 1
			End If
		Next
		PF3
		transmit
		panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes
	End IF
End If

Back_to_SELF
MAXIS_footer_month = right("00" & DatePart("m", SU_date), 2)
MAXIS_footer_year = right(DatePart ("yyyy", SU_date), 2)

If new_imig_checkbox = checked Then							'FSS CATEGORY - CLIENT IS A NEWLY ARRIVED IMMIGRANT
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 										'This variable determines which case notes will happen later
	new_imig_tikl = TRUE
	If new_imig_end_date = "" Then new_imig_end_date = DateAdd("m", 6, SU_date)
	new_imig_tikl_date = new_imig_end_date
 	fss_category_list = fss_category_list & "; Newly Arrived Immigrant"			'Creates a list of all the categories for case notes
	Do														'Go to IMIG
		Call Navigate_to_MAXIS_screen ("STAT", "IMIG")
		EMReadScreen nav_check, 4, 2, 49
	Loop until nav_check = "IMIG"
	EMWriteScreen ref_number, 20, 76
	transmit
	EMReadScreen numb_of_panels, 1, 2, 78
	IF numb_of_panels = "0" Then 							'If there is no IMIG panel the script stops because this can't be done
		script_end_procedure("ERROR: No IMIG Panel exists for this person. This coding cannot be completed for someone without an IMIG panel. The script will now end.")
	Else
		PF9
	End IF
	EMWriteScreen "Y", 18, 56								'Edit the panel amd add Yes code to the ELL code
	transmit
	panels_updated = panels_updated & "IMIG for Memb " & ref_number & " & "			'Creates a list of panels updated for case notes
End If

If IsDate(smc_diagnosis_date) = TRUE Then
	MAXIS_footer_month = right ("00" & DatePart ("m",smc_diagnosis_date), 2)
	MAXIS_footer_year = right (DatePart("yyyy", smc_diagnosis_date), 2)
End IF

If Special_medical_checkbox = checked Then 							'FSS CATEGORY - SOMEONE IN THE HOUSEHOLD MEETS SPECIAL MEDICAL CRITERIA
	STATS_counter = STATS_counter + 1
	fvw_only = FALSE 												'This variable determines which case notes will happen later
	smc_tikl = TRUE
	If smc_end_date = "" Then smc_end_date = DateAdd("m", 6, SU_date)
	smc_tikl_date = smc_end_date
	fss_category_list = fss_category_list & "; Special Medical Criteria"		'Creates a list of all the categories for case notes
	Do 																'Go to EMPS
		Call Navigate_to_MAXIS_screen ("STAT", "EMPS")
		EMReadScreen nav_check, 4, 2, 50
	Loop until nav_check = "EMPS"
	EMWriteScreen ref_number, 20, 76
	transmit
	PF9
	Select Case medical_criteria									'Logic to write the code based on the dropdown from the previous dialog
	Case "1 - Home-Health/Waiver Services"
		EMWriteScreen "1", 8, 76
	Case "2 - Child who meets SED Criteria"
		EMWriteScreen "2", 8, 76
	Case "3 - other Adult who meets SPMI"
		EMWriteScreen "3", 8, 76
	End Select
	transmit
	panels_updated = panels_updated & "EMPS for Memb " & ref_number & " & "

	next_MAXIS_month = CM_plus_1_mo & "/" & CM_plus_1_yr			'Going to STAT TIME to code in any banked months from specidal madical criteria
	TANF_banked_month = MAXIS_footer_month
	TANF_banked_year = MAXIS_footer_year
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "TIME")
		EMReadScreen nav_check, 4, 2, 46
	Loop until nav_check = "TIME"
	EMWriteScreen ref_number, 20, 76
	transmit
	PF9
	Do 																'Setting the locations of the fields on TIME
		If TANF_banked_month = "01" Then smc_month_col = 15
		If TANF_banked_month = "02" Then smc_month_col = 20
		If TANF_banked_month = "03" Then smc_month_col = 25
		If TANF_banked_month = "04" Then smc_month_col = 30
		If TANF_banked_month = "05" Then smc_month_col = 35
		If TANF_banked_month = "06" Then smc_month_col = 40
		If TANF_banked_month = "07" Then smc_month_col = 45
		If TANF_banked_month = "08" Then smc_month_col = 50
		If TANF_banked_month = "09" Then smc_month_col = 55
		If TANF_banked_month = "10" Then smc_month_col = 60
		If TANF_banked_month = "11" Then smc_month_col = 65
		If TANF_banked_month = "12" Then smc_month_col = 70
		For row = 5 to 16											'Looking in each of the field coordinates to find the month currently looking at
			EMReadScreen find_year, 2, row, 11
			If TANF_banked_year = find_year Then
				smc_month_row = row
				Exit For
			End If
		Next
		EMReadScreen is_counted, 2, smc_month_row, smc_month_col
		If is_counted = "SF" OR is_counted = "WF" Then 					'Fed funded months to change
			EMWriteScreen "FM", smc_month_row, smc_month_col
			tanf_banked_months_coded = tanf_banked_months_coded + 1
			banked_months_changed = banked_months_changed & " & " & TANF_banked_month & "/" & TANF_banked_year	'Create a list of months changed to case note
		ElseIF is_counted = "SS" OR is_counted = "WS" Then 				'State funded months to change
			EMWriteScreen "SM", smc_month_row, smc_month_col
			tanf_banked_months_coded = tanf_banked_months_coded + 1
			banked_months_changed = banked_months_changed & " & " & TANF_banked_month & "/" & TANF_banked_year	'Create a list of months changed to case note
		End If
		Call month_change(1, TANF_banked_month, TANF_banked_year, TANF_banked_month, TANF_banked_year)	'go to the next month
	Loop until TANF_banked_month & "/" & TANF_banked_year = next_MAXIS_month		'Loop until we get to cm + 1
	transmit
	panels_updated = panels_updated & "TIME for Memb " & ref_number & " & "
	If banked_months_changed <> "" Then banked_months_changed = right(banked_months_changed, len(banked_months_changed)-3)	'Formatting the list for case noting
End If

inhibiting_error = FALSE
month_to_start = CM_mo
year_to_start = CM_yr

'IF any of the code above had selections that stopped it from going forward, this will keep the script from case noting or TIKLing when no action was taken
IF fss_category_list = "" Then script_end_procedure("ERROR: There were no FSS Codes that could be updated with the information provided. Review your case.")

Do															'Going to STAT SUMM and sending it through background for getting MFIP results.
	Call Navigate_to_MAXIS_screen ("STAT", "SUMM")
	EMReadScreen nav_check, 4, 2, 46
Loop until nav_check = "SUMM"
EMWriteScreen "BGTX", 20, 71
transmit
Call date_array_generator (month_to_start, year_to_start, stat_date_array)	'creating an array of months of result to check
For Each version in stat_date_array											'Looking if a REVW is pending
	MAXIS_footer_month = right("00" & datepart("m", version), 2)
	MAXIS_footer_year = right(datepart("yyyy", version), 2)
	Do
		Call Navigate_to_MAXIS_screen ("STAT", "REVW")
		EMReadScreen revw_panel_check, 4, 2, 46
	Loop until revw_panel_check = "REVW"
	If er_due <> TRUE Then
		EMReadScreen er_code, 1, 7, 40
		Select Case er_code
		Case "_", "A"			'If this is _ or A this case will not be held up for processing a review
			er_due = FALSE
		Case "I", "N"			'If it is coded like this then case cannot be approved due to an ER due
			er_due = TRUE
			er_due_month = MAXIS_footer_month & "/" & MAXIS_footer_year
		End Select
	End If
	If mont_due <> TRUE Then 			'Looking for HRF due - because no approval can be done
		Call Navigate_to_MAXIS_screen ("STAT", "MONT")
		EMReadScreen mont_code, 1, 11, 43
		Select Case mont_code
		Case "_", "A"
			mont_due = FALSE
		Case "I", "N"
			mont_due = TRUE
			mont_due_month = MAXIS_footer_month & "/" & MAXIS_footer_year
		End Select
	End If
	Back_to_SELF
Next

'Adding verbiage to reasons not approved if HRF or ER are due
If er_due = TRUE Then notes_not_approved = notes_not_approved & "ER due for " & er_due_month & "; "
IF mont_due = TRUE Then notes_not_approved = notes_not_approved & "HRF due for " & mont_due_month & "; "

month_to_start = CM_mo
year_to_start = CM_yr

'Reformatting all the lists created
fss_category_list = right(fss_category_list, len(fss_category_list) - 1) & ""
If panels_updated <> "" Then panels_updated = left(panels_updated, len(panels_updated)-3)
If panels_reviewed <> "" Then panels_reviewed = left(panels_reviewed, len(panels_reviewed)-3)
If other_notes <> "" THEN other_notes = left(other_notes, len(other_notes)-1) & ""
If notes_not_approved <> "" Then notes_not_approved = left (notes_not_approved, len(notes_not_approved)-2)
Call Read_MFIP_Results(month_to_start, year_to_start, MFIP_results)		'Getting the detail about MFIP results

'Runs the final dialog
Do
	Do
		err_msg = ""
		Dialog FSS_final_dialog
		Cancel_confirmation
		MAXIS_dialog_navigation
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "Sign your case note!"
		If results_approved_checkbox = unchecked AND not_approved_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "You must indicate if you approved the new MFIP results or not."
		IF results_approved_checkbox = checked AND not_approved_checkbox = checked Then err_msg = err_msg & vbNewLine & "You must pick if you have approved the new MFIP results or not - it cannot be both."
		IF not_approved_checkbox = checked AND notes_not_approved = "" Then err_msg = err_msg & vbNewLine & "If you did not approve the new MFIP results, you must explain why the approval is not being done."
		If ButtonPressed = CASE_BGTX_button Then
			err_msg = err_msg & "new results needed"
			Call Read_MFIP_Results(month_to_start, year_to_start, MFIP_results)
		End IF
		If err_msg <> "" AND ButtonPressed <> CASE_BGTX_button Then MsgBox "** Resolve to continue **" & vbNewLine & vbNewLine & err_msg
	Loop until ButtonPressed = OK AND err_msg = ""
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'Writing a TIKL if needed
'ILL or INCAPACITATED
ill_incap_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case." & client_name & " is listed as being ill or incapacitated."
ill_incap_tikl_fail    = "Script could not write a TIKL for the end of ill/incapacitated FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed(ill_incap_tikl, ill_incap_tikl_date, ill_incap_TIKL_message, ill_incap_tikl_fail)

'CARE OF ILL OF INCAPACITATED HH MEMB
care_of_ill_incap_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case." & client_name & " is listed as caring  for ill/incap HH Member."
care_of_ill_incap_tikl_fail    = "Script could not write a TIKL for the end of care of ill/incapacitated FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed(care_of_ill_incap_tikl, care_of_ill_incap_tikl_date, care_of_ill_incap_TIKL_message, care_of_ill_incap_tikl_fail)

'CHILD UNDER ONE
child_under_one_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case. Child under one year exemption to end this month. Case needs to be sent through background."
child_under_one_tikl_fail    = "Script could not write a TIKL for the end of child under one FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed (child_under_one_tikl, child_under_one_tikl_date, child_under_one_TIKL_message, child_under_one_tikl_fail)

'FAMILY VIOLENCE WAIVER
fvw_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case. Case is coded as meeting a Family Violence Waiver."
fvw_tikl_fail    = "Script could not write a TIKL for the Family Violence Wavier FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed(fvw_tikl, fvw_tikl_date, fvw_TIKL_message, fvw_tikl_fail)

'SPECIAL MEDICAL CRITERIA
smc_tikl_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case. Case is coded for Specidal Medical Criteria and needs assesment."
smc_tikl_fail         = "Script could not write a TIKL for the end of Special Medical Criteria FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed (smc_tikl, smc_tikl_date, smc_tikl_TIKL_message, smc_tikl_fail)

'HARD TO EMPLOY
hard_to_employ_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case. Case is coded that " & client_name & " meets a hard to employ category and needs review."
hard_to_employ_tikl_fail    =  "Script could not write a TIKL for the end of Hard to Employ FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed (hard_to_employ_tikl, hard_to_employ_tikl_date, hard_to_employ_TIKL_message, hard_to_employ_tikl_fail)

'PENDING SSI OR RSDI
ssi_pending_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case." & client_name & " is pending SSI/RSDI - review status of application."
ssi_pending_tikl_fail    = "Script could not write a TIKL for the end of SSI/RSDI Pending FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed(ssi_pending_tikl, ssi_pending_tikl_date, ssi_pending_TIKL_message, ssi_pending_tikl_fail)

'NEWLY ARRIVED IMMIGRANT\
new_imig_tikl_TIKL_message = "^^^ FSS Expired ^^^ MFIP case is codded for FSS. Review case." & client_name & " is coded as a newly arrived immigrant enrolled in ESL Skills Training - review status."
new_imig_tikl_fail         = "Script could not write a TIKL for the end of Newly Arrived Immigrant in ESL FSS code. You will need to set the TIKL manually."
Call Write_TIKL_if_needed(new_imig_tikl, new_imig_tikl_date, new_imig_tikl_TIKL_message, new_imig_tikl_fail)

'Case note for the Family Violence Waiver
IF fam_violence_checkbox = checked Then
	CALL start_a_blank_CASE_NOTE
	IF new_fvw = TRUE Then 		'For new waivers being added
		CALL write_variable_in_CASE_NOTE ("***** DOMESTIC VIOLENCE WAIVER *****")
	ElseIF new_fvw = FALSE Then		'For renewal of waiver
		CALL write_variable_in_CASE_NOTE ("***** DVW RENEWED *****")
	End IF
	CALL write_bullet_and_variable_in_CASE_NOTE ("Effective Date", fvw_start_date)
	CALL write_bullet_and_variable_in_CASE_NOTE ("End Date", fvw_end_date)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Worker", es_worker)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Agency", es_agency)
	CALL write_variable_in_CASE_NOTE ("* All documentation needed for the waiver is with Employment Services, including advocate information and review details.")
	CALL write_bullet_and_variable_in_CASE_NOTE ("Months Changed due to Waiver", counted_months_changed)
	CALL write_bullet_and_variable_in_CASE_NOTE ("TANF Months Used", tanf_used)
	CALL write_bullet_and_variable_in_CASE_NOTE ("Extension Months Used", ext_tanf_used)
	IF fvw_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Family Violence at " & fvw_tikl_date)
	CALL write_variable_in_CASE_NOTE ("---")
	IF results_approved_checkbox = checked Then CALL write_bullet_and_variable_in_CASE_NOTE ("MFIP Results Approved", MFIP_results)
	IF not_approved_checkbox = checked Then Call write_bullet_and_variable_in_CASE_NOTE ("New MFIP NOT Approved Due To", notes_not_approved)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE (worker_signature)
End IF

'Case note for any other FSS category
If fvw_only = FALSE Then
	CALL start_a_blank_CASE_NOTE
	IF new_fss = TRUE Then 		'For new FSS cases
		CALL write_variable_in_CASE_NOTE ("**** FSS ELIGIBLE ****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Approved change to state funding effective", month_to_start & "/" & year_to_start)
	ElseIF new_fss = FALSE Then  		'For cases already coded as FSS to extend
		CALL write_variable_in_CASE_NOTE ("**** FSS CONTINUES ****")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Approved continued state funding effective", month_to_start & "/" & year_to_start)
	End IF
	CALL write_bullet_and_variable_in_CASE_NOTE ("Eligibility of Category", fss_category_list)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Worker", es_worker)
	CALL write_bullet_and_variable_in_CASE_NOTE ("ES Agency", es_agency)
	If ill_incap_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- Caregiver Ill or Incap ---")
		IF ill_incap_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of clt Ill/Incap is with Employment Services.")
		IF ill_incap_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of clt Ill/Incap is with Financial Case File.")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Ill/Incap Start Date", ill_incap_start_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Ill/Incap End Date", ill_incap_end_date)
		IF ill_incap_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Ill/Incap at " & ill_incap_tikl_date)
	End If
	If care_of_ill_Incap_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- Care of an Ill or Incap HH Memb ---")
		IF rel_care_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Ill/Incap HH Member is with Employment Services.")
		IF rel_care_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Ill/Incap HH Member is with Financial Case File.")
		CALL write_variable_in_CASE_NOTE ("* Caregiver is required in the home to care for " & disa_HH_memb)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Relative Care Start Date", rel_care_start_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Relative Care End Date", rel_care_end_date)
		IF care_of_ill_incap_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for care of Ill/Incap HH member at " & care_of_ill_incap_tikl_date)
	End If
	If iq_test_checkbox = checked OR learning_disabled_checkbox = checked OR mentally_ill_checkbox = checked OR dev_delayed_checkbox = checked OR unemployable_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- Caregiver meets Hard to Employ Category ---")
		IF unemployable_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Unemployability is with Employment Services.")
		IF unemployable_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Unemployability is with Financial Case File.")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Hard to Employ Start Date", unemployable_start_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Hard to Employ End Date", unemployable_end_date)
		IF hard_to_employ_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Hard to Employ Category at " & hard_to_employ_tikl_date)
	End IF
	If ssi_pending_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- Caregiver is Pending SSI/RSDI ---")
		IF ssa_app_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Application for SSI/RSDI is with Employment Services.")
		IF ssa_app_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Application for SSI/RSDI is with Financial Case File.")
		CALL write_bullet_and_variable_in_CASE_NOTE ("SSA App Date", ssa_app_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("End Date of SSA App category", ssa_end_date)
		IF ssi_pending_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Pending SSI/RSDI at " & ssi_pending_tikl_date)
	End If
	If child_under_one_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- Child Under 12 Months Exemption ---")
		IF child_under_1_at_es = checked Then CALL write_variable_in_CASE_NOTE ("* Request to take the Child Under 12 Months exemption was made to ES Worker.")
		IF child_under_1_at_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Request to take the Child Under 12 Months exemption was made to FW.")
		IF TIKL_verified = TRUE Then CALL write_variable_in_CASE_NOTE ("* TIKL set to end the exemption and do a new MFIP approval when months are all used.")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Months coded for Child < 12 Months Exemption", Join(new_exemption_months_array, ", "))
		IF child_under_one_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Child Under One at " & child_under_one_tikl_date)
	End IF
	If new_imig_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- Newly Arrived Immigrant ---")
		CALL write_variable_in_CASE_NOTE ("* Documentation of particilation with ELL Classes and SPL is with Employment Services.")
		IF new_imig_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Newly Arrived Immigrant at " & new_imig_tikl_date)
	End If
	If Special_medical_checkbox = checked Then
		CALL write_variable_in_CASE_NOTE ("--- HH Member Meets Special Medical Criteria ---")
		IF smc_docs_with_es = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Special Medical Criteria is with Employment Services.")
		IF smc_docs_with_fas = checked Then CALL write_variable_in_CASE_NOTE ("* Documentation of Special Medical Criteria is with Financial Case File.")
		CALL write_variable_in_CASE_NOTE ("* Special Medical Criteria for Memb " & smc_hh_memb & " for " & medical_criteria)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Date of Diagnosis", smc_diagnosis_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Banked Months Changed on TIME", banked_months_changed)
		IF smc_tikl = TRUE Then Call write_variable_in_CASE_NOTE ("* TIKL set to review FSS for Special Medical Criteria at " & smc_tikl_date)
	End IF
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_bullet_and_variable_in_CASE_NOTE ("STAT Panels Updated", panels_updated)
	CALL write_bullet_and_variable_in_CASE_NOTE ("STAT Panels Reviewed", panels_reviewed)
	CALL write_bullet_and_variable_in_CASE_NOTE ("NOTES", other_notes)
	CALL write_variable_in_CASE_NOTE ("---")
	IF results_approved_checkbox = checked Then CALL write_bullet_and_variable_in_CASE_NOTE ("MFIP Results Approved", MFIP_results)
	IF not_approved_checkbox = checked Then Call write_bullet_and_variable_in_CASE_NOTE ("New MFIP NOT Approved Due To", notes_not_approved)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE (worker_signature)
End If


STATS_counter = STATS_counter - 1
script_end_procedure("Success! STAT updated and Case Note entered for FSS.")
