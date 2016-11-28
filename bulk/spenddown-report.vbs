'Required for statistical purposes==========================================================================================
name_of_script = "BULK - SPENDDOWN REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 72                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
'END OF stats block==============================================================================================

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

'This function is used to grab all active X numbers according to the supervisor X number(s) inputted
FUNCTION create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
	'Getting to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")


	'Sorting by supervisor
	PF5
	PF5


	'Reseting array_name
	array_name = ""


	'Splitting the list of inputted supervisors...
	supervisor_array = replace(supervisor_array, " ", "")
	supervisor_array = split(supervisor_array, ",")
	FOR EACH unit_supervisor IN supervisor_array
		IF unit_supervisor <> "" THEN
			'Entering the supervisor number and sending a transmit
			CALL write_value_and_transmit(unit_supervisor, 21, 12)


			MAXIS_row = 7
			DO
				EMReadScreen worker_ID, 8, MAXIS_row, 5
				worker_ID = trim(worker_ID)
				IF worker_ID = "" THEN EXIT DO
				array_name = trim(array_name & " " & worker_ID)
				MAXIS_row = MAXIS_row + 1
				IF MAXIS_row = 19 THEN
					PF8
					EMReadScreen end_check, 9, 24,14
					If end_check = "LAST PAGE" Then Exit Do
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
END FUNCTION



'DIALOGS----------------------------------------------------------------------
BeginDialog find_spenddowns_month_spec_dialog, 0, 0, 221, 180, "Pull REPT data into Excel dialog"
  EditBox 85, 20, 130, 15, worker_number
  DropListBox 125, 85, 80, 45, "ALL"+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", revw_month_list
  CheckBox 5, 105, 150, 10, "Check here to have the script check MMIS", MMIS_checkbox
  CheckBox 5, 120, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 110, 160, 50, 15
    CancelButton 165, 160, 50, 15
  Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 60, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
  Text 5, 90, 120, 10, "Only pull cases with next review in:"
  Text 5, 135, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog


'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

'Setting this variable to determine a filter later
one_month_only = FALSE
'defining current footer month so the script doesn't go to old things
MAXIS_footer_month = right("00" & datepart("m", date), 2)
MAXIS_footer_year = right("00" & datepart("yyyy", date), 2)

'Shows dialog
Dialog find_spenddowns_month_spec_dialog
If buttonpressed = cancel then stopscript

'Sets the script up to only pull cases for certain months if selected from the dialog
If revw_month_list <> "ALL" AND revw_month_list <> "" Then
	one_month_only = TRUE 			'If any month is selected the script needs to filter
	Select Case revw_month_list
		Case "January"
			month_selected = 1
		Case "February"
			month_selected = 2
		Case "March"
			month_selected = 3
		Case "April"
			month_selected = 4
		Case "May"
			month_selected = 5
		Case "June"
			month_selected = 6
		Case "July"
			month_selected = 7
		Case "August"
			month_selected = 8
		Case "September"
			month_selected = 9
		Case "October"
			month_selected = 10
		Case "November"
			month_selected = 11
		Case "December"
			month_selected = 12
	End Select
End If

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else		'If worker numbers are litsted - this will create an array of workers to check
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		x1_number = trim(ucase(x1_number))					'Formatting the x numbers so there are no errors
		Call navigate_to_MAXIS_screen ("REPT", "USER")		'This part will check to see if the x number entered is a supervisor of anyone
		PF5
		PF5
		EMWriteScreen x1_number, 21, 12
		transmit
		EMReadScreen sup_id_check, 7, 7, 5					'This is the spot where the first person is listed under this supervisor
		IF sup_id_check <> "       " Then 					'If this frist one is not blank then this person is a supervisor
			supervisor_array = trim(supervisor_array & " " & x1_number)		'The script will add this x number to a list of supervisors
		Else
			If worker_array = "" then						'Otherwise this x number is added to a list of workers to run the script on
				worker_array = trim(x1_number)
			Else
				worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
			End if
		End If
		PF3
	Next

	If supervisor_array <> "" Then 				'If there are any x numbers identified as a supervisor, the script will run the function above
		Call create_array_of_all_active_x_numbers_by_supervisor (more_workers_array, supervisor_array)
		workers_to_add = join(more_workers_array, ", ")
		If worker_array = "" then				'Adding all x numbers listed under the supervisor to the worker array
			worker_array = workers_to_add
		Else
			worker_array = worker_array & ", " & trim(ucase(workers_to_add))
		End if
	End If

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting up constants for ease of reading the array
Const wrk_num   = 0
Const case_num  = 1
Const next_revw = 2
Const clt_name  = 3
Const ref_numb  = 4
Const clt_pmi   = 5
Const hc_type   = 6
Const mobl_spdn = 7
Const spd_pd    = 8
Const hc_excess = 9
Const mmis_spdn = 10
Const add_xcl   = 11

'Setting up the arrays to be dynamic
Dim clts_with_spdwn_array()
ReDim clts_with_spdwn_array (3, 0)

Dim spenddown_error_array ()
ReDim spenddown_error_array (12, 0)

'Setting the variable for what's to come
excel_row = 2
hc_clt = 0

'Getting all the cases with HC active for each worker
For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")	'going to rept actv for each worker
	EMWriteScreen worker, 21, 13
	transmit
	EMReadScreen user_worker, 7, 21, 71
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7

			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21			'Reading client name
				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
				EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
				If HC_status = "A" then
					If one_month_only = TRUE Then 						'If user has selected to only get cases with a certain reveiw month
						If trim(next_revw_date) = "" Then
							case_error = MsgBox ("Case " & MAXIS_case_number & " does not have a review listed, please check that STAT is coded correctly for this case." & vbNewLine & vbNewLine & "This case will not be added to the report, you should check for a spenddown manually.", vbAlert, "No Review Date")
						Else
							revw_month = abs(left(next_revw_date, 2))
							If revw_month = month_selected Then 			'Compares the review month to the variable defined above in the Select Case
								ReDim Preserve clts_with_spdwn_array (3, hc_clt)		'Adds information about case with active HC to an array
								clts_with_spdwn_array(wrk_num, hc_clt)   = worker
								clts_with_spdwn_array(case_num, hc_clt)  = MAXIS_case_number
								clts_with_spdwn_array(next_revw, hc_clt) = next_revw_date
								hc_clt = hc_clt + 1
							End If
						End If
					Else
						ReDim Preserve clts_with_spdwn_array (3, hc_clt)			'Adds information about case with active HC to an array
						clts_with_spdwn_array(wrk_num, hc_clt)   = worker
						clts_with_spdwn_array(case_num, hc_clt)  = MAXIS_case_number
						clts_with_spdwn_array(next_revw, hc_clt) = next_revw_date
						hc_clt = hc_clt + 1
					End If
				End If


				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

spd_case = 0

'The script will now look in each case at MOBL to identify clients that have spenddown listed on MOBL'\
For hc_case = 0 to UBound(clts_with_spdwn_array, 2)
	MAXIS_case_number = clts_with_spdwn_array(case_num, hc_case)		'defining case number for functions to use
	Call navigate_to_MAXIS_screen ("ELIG", "HC")						'Goes to ELIG HC
	row = 8
	Do										'Looks at each row in HC Elig to find the first MA span
		EMReadScreen prog, 2, row, 28
		If prog = "MA" Then
			EMWriteScreen "X", row, 26		'Goes into it
			transmit
			Exit Do
		End if
		row = row + 1
	Loop until row = 20
	If row <> 20 Then 						'Once in the span, opens MOBL
		EMWriteScreen "X", 18, 3
		transmit
		Do
			EMReadScreen MOBL_check, 4, 3, 49
			If MOBL_check <> "MOBL" Then
				row = row + 1
				PF3
				EMReadScreen prog, 2, row, 28
				If prog = "MA" Then
					EMWriteScreen "X", row, 26		'Goes into it
					transmit
					EMWriteScreen "X", 18, 3
					transmit
				End if
			End If
		Loop until row = 20 OR MOBL_check = "MOBL"
		row = 6
		Do									'reads each line on MOBL and saves the clt information for any client that has a spenddown indicated on MOBL
			EMReadScreen spd_type, 20, row, 39
			spd_type = trim(spd_type)
			If spd_type = "" Then Exit Do			'Leaves the do loop once a blank line is found
			If spd_type <> "NO SPENDDOWN" Then 		'Anything other than this indicates MAXIS thinks there is a spenddown
				EMReadScreen reference, 2, row, 6
				EMReadScreen period, 13, row, 61
				EMReadScreen cname, 21, row, 10
				cname = trim(cname)
				If cname = "" Then EMReadScreen cname, 21, row - 1, 10
				cname = trim(cname)

				ReDim Preserve spenddown_error_array (12, spd_case)			'Adding any client with a spenddown to a new array

				spenddown_error_array (wrk_num,   spd_case) = clts_with_spdwn_array(wrk_num, hc_case)
				spenddown_error_array (case_num,  spd_case) = MAXIS_case_number
				spenddown_error_array (next_revw, spd_case) = replace(clts_with_spdwn_array(next_revw, hc_case), " ", "/")
				spenddown_error_array (clt_name,  spd_case) = cname
				spenddown_error_array (ref_numb,  spd_case) = reference
				spenddown_error_array (mobl_spdn, spd_case) = spd_type
				spenddown_error_array (spd_pd,    spd_case) = period

				spd_case = spd_case + 1

			End If
			row = row + 1
		Loop until row = 19
	End If
Next

'This bit will look to see if there are any cases that have a possible spenddown.
'Occasionally the criteria selected produce no cases and this explains this to the user.
If UBound(spenddown_error_array, 2) = 0 AND spenddown_error_array(case_num, 0) = "" Then
	all_workers = Join(worker_array, ", ")
	If one_month_only = True Then
		selected_time = " for the month of " & revw_month_list & "."
	Else
		selected_time = "."
	End If
	end_msg = "Success! The script has completed!" & vbNewLine & "NO SPENDDOWNS FOUND!" & vbNewLine & vbNewLine &_
	          "The script has checked REPT/ACTV for the case loads under worker number(s) " & all_workers & selected_time & vbNewLine &_
			  "None of the active HC cases have a spenddown indicated on MOBL." & vbNewLine & vbNewLine &_
			  "No report will be generated, the script has completed."
	script_end_procedure(end_msg)
End If

'Gathering additional information about each client with a spenddown indicated
For spd_case = 0 to UBound(spenddown_error_array, 2)
	spd_amt = 0			'Reset the variable for each run
	MAXIS_case_number = spenddown_error_array(case_num, spd_case)				'Setting the case number for global functions
	spenddown_error_array(add_xcl, spd_case) = TRUE
	Call navigate_to_MAXIS_screen ("CASE", "PERS")								'Confirming clt is active HC this month
	row = 9
	Do
		EMReadScreen person, 2, row, 3
		If person = spenddown_error_array(ref_numb, spd_case) Then
			EMReadScreen hc_stat, 1, row, 61
			If hc_stat = "I" OR hc_stat = "D" Then spenddown_error_array(add_xcl, spd_case) = FALSE 	'If not, case will not be added to report
			Exit Do
		Else
			row = row + 1
			If row = 18 Then
				EMReadScreen next_page, 7, row, 3
				If next_page = "More: +" Then
					PF8
					row = 9
				End If
			End If
		End If
	Loop until row = 18
	IF spenddown_error_array(add_xcl, spd_case) = TRUE Then 						'If clt is actve HC
		STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
		Call navigate_to_MAXIS_screen ("ELIG", "HC")								'Need a closer look at HC
		row = 8
		Do
			EMReadScreen person, 2, row, 3									'Finding the correct person on HC ELIG
			If person = spenddown_error_array(ref_numb, spd_case) Then
				Do
					EMReadScreen prog, 2, row, 28							'Find the line that has this persons MA listed on it /NOT QMB etc
					If prog = "MA" Then
						counter = 1
						Do 													'Here this will find if this version is approved
							EMReadScreen app_indc, 5, row, 68
							app_indc = trim(app_indc)
							If app_indc = "APP" Then						'If approved, open this version
								Call write_value_and_transmit("x", row, 26)
								Exit Do
							End If
							EMReadScreen this_version, 2, row, 58
							If this_version = "00" Then
								EMReadScreen elig_month, 2, 20, 56			'If the earliest version in this month was not approved then it goes to the previous month
								If elig_month <> "01" Then
									last_month = right("00" & (abs(elig_month)-1), 2)
									EMWriteScreen last_month, 20, 56
									transmit
								Else
									last_month = "12"
									EMReadScreen elig_year, 2, 20, 59
									last_year = right("00" & (abs(elig_year)-1), 2)
									EMWriteScreen last_month, 20, 56
									EMWriteScreen last_year, 20, 59
								End If
								counter = counter + 1
							ElseIf this_version <> "01" Then 					'Checking to see if there is a pervious version listed in this month
								prev_verision = right("00" & (abs(this_version)-1), 2)	'If so, it will go to the previous version
								EMWriteScreen prev_verision, row, 58
								transmit
							Else
								EMReadScreen elig_month, 2, 20, 56			'If the earliest version in this month was not approved then it goes to the previous month
								If elig_month <> "01" Then
									last_month = right("00" & (abs(elig_month)-1), 2)
									EMWriteScreen last_month, 20, 56
									transmit
								Else
									last_month = "12"
									EMReadScreen elig_year, 2, 20, 59
									last_year = right("00" & (abs(elig_year)-1), 2)
									EMWriteScreen last_month, 20, 56
									EMWriteScreen last_year, 20, 59
								End If
								counter = counter + 1
							End If
						Loop until counter = 6				'Only looks at 6 months
					ELSE					'If this person could not be found then the report will list no version was found
						row = row + 1
						EMReadScreen person, 2, row, 3
						If person <> "  "  Then
							spenddown_error_array(hc_type, spd_case) = "NO HC VERSION"
							Exit Do
						End If
					End If
				Loop Until row = 20
				Exit Do
			Else
				row = row + 1
			End If
		Loop until row = 20
		EMReadScreen bsum_check, 4, 3, 57		'Confirming that HC Elig has been opened for this person
		If bsum_check = "BSUM" Then
			col = 19
			Do									'Finding the current month in elig to get the current elig type
				EMReadScreen span_month, 2, 6, col
				If span_month = MAXIS_footer_month Then		'reading the ELIG TYPE
					EMReadScreen pers_type, 2, 12, col - 2
					EMReadScreen std, 1, 12, col + 3
					EMReadScreen meth, 1, 13, col + 2
					Exit Do
				End If
				col = col + 11
				If col = 85 Then 		'If this month was not found then it reads the LAST elig type in elig
					EMReadScreen pers_type, 2, 12, 72		'ONLY saves this information if an actual elig type was found
					If pers_type <> "11" AND pers_type <> "09" AND pers_type <> "PX" AND pers_type <> "PC" AND pers_type <> "CB" AND pers_type <> "CK" AND pers_type <> "CX" AND pers_type <> "CM" AND pers_type <> "AA" AND pers_type <> "AX" AND pers_type <> "BT" AND pers_type <> "DT" AND pers_type <> "15" AND pers_type <> "16" AND pers_type <> "DC" AND pers_type <> "EX" AND pers_type <> "DX" AND pers_type <> "DP" AND pers_type <> "BC" AND pers_type <> "RM" AND pers_type <> "10" AND pers_type <> "25" Then
						pers_type = ""
					Else
						EMReadScreen std, 1, 12, 77
						EMReadScreen meth, 1, 13, 76
					End If
				End If
			Loop until col = 85
			If pers_type = "" Then 				'Setting the elig type to readable format
				spenddown_error_array(hc_type, spd_case) = "ELIG Type Not Found"
			Else
				spenddown_error_array(hc_type, spd_case) = pers_type & "-" & std & " Method: " & meth
				pers_type = ""
				std = ""
				meth = ""
			End If
			spd_amt = 0
			col = 18
			Do 				'This will gather the 6 month standard AND the budgeted income to calculate the HC overage
				EMReadScreen month_net_inc, 8, 15, col
				EMReadScreen month_std_inc, 8, 16, col
				month_net_inc = trim(month_net_inc)
				If month_net_inc = "" Then month_net_inc = 0
				month_std_inc = trim(month_std_inc)
				If month_std_inc = "" Then month_std_inc = 0
				tot_net_inc = tot_net_inc + abs(month_net_inc)
				tot_std_inc = tot_std_inc + abs(trim(month_std_inc))
				col = col + 11
			Loop until col = 84
			spd_amt =  tot_net_inc - tot_std_inc
			If spd_amt < 0 Then spd_amt = 0
			spenddown_error_array(hc_excess, spd_case) = spd_amt
			'NOTE that Cert Period Amount popup was NOT used as it appears to change what is listed on MOBL if the spenddown was in error
			'We do not want bulk reports to make alterations to cases without worker review and approval
		End If

		'Goes to get PMI
		Call navigate_to_MAXIS_screen ("STAT", "MEMB")
		EMWriteScreen spenddown_error_array(ref_numb, spd_case), 20, 76
		transmit
		EMReadScreen pmi, 8, 4, 46
		spenddown_error_array(clt_pmi, spd_case) = right("00000000" & replace(pmi, "_", ""), 8)
	End If
	back_to_self
Next

If MMIS_checkbox = checked Then
	'Now it will look for MMIS on both screens, and enter into it..
	attn
	EMReadScreen MMIS_A_check, 7, 15, 15
	If MMIS_A_check = "RUNNING" then
		EMSendKey "10"
		transmit
	Else
		attn
		EMConnect "B"
		attn
		EMReadScreen MMIS_B_check, 7, 15, 15
		If MMIS_B_check <> "RUNNING" then
			MMIS_checkbox = unchecked
			script_continue = MsgBox ("MMIS does not appear to be running." & vbNewLine & "Do you wish to have the report without the MMIS Spenddown Indicator checked?", vbYesNo + vbQuestion, "MMIS not running")
			IF script_continue = vbNo Then script_end_procedure ("Script has ended with no report generated. To have MMIS information gathered, be sure to have MMIS running and not be passworded out.")
		Else
			EMSendKey "10"
			transmit
		End if
	End if
End If

If MMIS_checkbox = checked Then
	EMFocus 'Bringing window focus to the second screen if needed.

	'Sending MMIS back to the beginning screen and checking for a password prompt
	Do
		PF6
		EMReadScreen password_prompt, 38, 2, 23
	  	IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then
		  	MMIS_checkbox = unchecked
		  	script_continue = MsgBox ("MMIS does not appear to be running." & vbNewLine & "Do you wish to have the report without the MMIS Spenddown Indicator checked?", vbYesNo + vbQuestion, "MMIS not running")
		  	IF script_continue = vbNo Then script_end_procedure ("Script has ended with no report generated. To have MMIS information gathered, be sure to have MMIS running and not be passworded out.")
			Exit Do
		End If
	  	EMReadScreen session_start, 18, 1, 7
	Loop until session_start = "SESSION TERMINATED"
End If

If MMIS_checkbox = checked Then
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themself into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit

	'Finding the right MMIS, if needed, by checking the header of the screen to see if it matches the security group selector
	EMReadScreen MMIS_security_group_check, 21, 1, 35
	If MMIS_security_group_check = "MMIS MAIN MENU - MAIN" then
		EMSendKey "x"
		transmit
	End if

	'Now it finds the recipient file application feature and selects it.
	row = 1
	col = 1
	EMSearch "RECIPIENT FILE APPLICATION", row, col
	EMWriteScreen "x", row, col - 3
	transmit

	For spd_case = 0 to UBound(spenddown_error_array, 2)			'Opens RELG for each client to get spenddown indicator
		indicator = ""
		'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
		EMWriteScreen "i", 2, 19
		EMWriteScreen spenddown_error_array(clt_pmi, spd_case), 4, 19	'Enters PMI
		transmit		'Goes to RSUM
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then 		'Confirms that we have moved past RKEY
			spenddown_error_array (mmis_spdn, spd_case) = "Not Found"
		Else
			EMWriteScreen "RELG", 1, 8		'Goes to RELG
			transmit
			row = 7
			Do 				'Finding the openended OR future close MA span
				EMReadscreen elig_end, 8, row, 36
				IF elig_end <> "99/99/99" Then after_now = DateDiff("d", date, elig_end)
				If elig_end = "99/99/99" OR after_now < 0 Then
					EMReadscreen prg, 2, row-1, 10
					IF prg = "MA" Then 			'Reads the spenddown indicator
						EMReadscreen indicator, 1, row + 1, 62
						Exit Do
					End If
				End If
				row = row + 4
			Loop until row = 23

			PF6
			EMWriteScreen "        ", 4, 19		'Blanking out the PMI for safety

			If indicator = "" Then 				'Setting the indicator to the array
				spenddown_error_array (mmis_spdn, spd_case) = "Not Found"
			Else
				spenddown_error_array (mmis_spdn, spd_case) = indicator
			End If
		End If

	Next
End If

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "REF NO"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "NAME"
objExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "PMI"
objExcel.Cells(1, 5).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "NEXT REVW DATE"
objExcel.Cells(1, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 7).Value = "ELIG TYPE"
objExcel.Cells(1, 7).Font.Bold = TRUE
ObjExcel.Cells(1, 8).Value = "SPDWN ON MOBL"
objExcel.Cells(1, 8).Font.Bold = TRUE
ObjExcel.Cells(1, 9).Value = "HC OVERAGE"
objExcel.Cells(1, 9).Font.Bold = TRUE
If MMIS_checkbox = checked Then
	ObjExcel.Cells(1, 10).Value = "MMIS SPDWN"
	objExcel.Cells(1, 10).Font.Bold = TRUE
End If

'Adding all client information to a spreadsheet for your viewing pleasure
For spd_case = 0 to UBound(spenddown_error_array, 2)
	If spenddown_error_array(add_xcl, spd_case) = TRUE Then
		ObjExcel.Cells(excel_row, 1).Value  = spenddown_error_array (wrk_num,   spd_case)
		ObjExcel.Cells(excel_row, 2).Value  = spenddown_error_array (case_num,  spd_case)
		ObjExcel.Cells(excel_row, 3).Value  = "Memb " & spenddown_error_array(ref_numb, spd_case)
		ObjExcel.Cells(excel_row, 4).Value  = spenddown_error_array (clt_name,  spd_case)
		ObjExcel.Cells(excel_row, 5).Value  = spenddown_error_array (clt_pmi,   spd_case)
		ObjExcel.Cells(excel_row, 6).Value  = spenddown_error_array (next_revw, spd_case)
		ObjExcel.Cells(excel_row, 7).Value  = spenddown_error_array (hc_type,   spd_case)
		ObjExcel.Cells(excel_row, 8).Value  = spenddown_error_array (mobl_spdn, spd_case) & " for " & spenddown_error_array(spd_pd, spd_case)
		ObjExcel.Cells(excel_row, 9).Value  = spenddown_error_array (hc_excess, spd_case)
		ObjExcel.Cells(excel_row, 10).Value = spenddown_error_array (mmis_spdn, spd_case)

		excel_row = excel_row + 1
	End If
Next

'Query date/time/runtime info
objExcel.Cells(1, 11).Font.Bold = TRUE
objExcel.Cells(2, 11).Font.Bold = TRUE
ObjExcel.Cells(1, 11).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 12).Value = now
ObjExcel.Cells(2, 11).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 12).Value = timer - query_start_time


'Autofitting columns
For col_to_autofit = 1 to 12
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! All cases for selected workers that appear to have a Spenddown indicated in MAXIS have been added to the Excel Spreadsheet.")
