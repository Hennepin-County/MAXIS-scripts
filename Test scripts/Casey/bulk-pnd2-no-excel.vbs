'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-PND2 LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

const worker_numb_const	= 0
const case_numb_const 	= 1
const case_name_const 	= 2
const appl_date_const 	= 3
const days_pending_const= 4
const snap_const 		= 5
const cash_const 		= 6
const cash_prog_const 	= 7
const cash_1_const 		= 8
const cash_1_prog_const = 9
const cash_2_const 		= 10
const cash_2_prog_const = 11
const hc_const 			= 12
const ea_const 			= 13
const grh_const 		= 14
const ive_const 		= 15
const ccap_const 		= 16
const pnd2_last_const 	= 20

Dim PND2_ARRAY()
ReDim PND2_ARRAY(pnd2_last_const, 0)


'THE SCRIPT-------------------------------------------------------------------------
'Gathering county code for multi-county...
get_county_code

worker_number = "X127EE2, X127EN5, X127EG4, X127ED8, X127EH8, X127EQ3, X127EQ2, X127ET9, X127ET8, X127ES8, X127ES3, X127EP6, X127EP7, X127EP8"

'Connects to BlueZone
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 286, 135, "Pull REPT data into Excel dialog"
  EditBox 135, 20, 145, 15, worker_number
  CheckBox 70, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 70, 80, 205, 10, "Check here to add last case note information to spreadsheet.", case_note_check
  CheckBox 10, 20, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 35, 40, 10, "Cash?", cash_check
  CheckBox 10, 50, 40, 10, "HC?", HC_check
  CheckBox 10, 65, 40, 10, "EA?", EA_check
  CheckBox 10, 80, 40, 10, "GRH?", GRH_check
  CheckBox 10, 95, 40, 10, "IV-E?", IVE_check
  CheckBox 10, 110, 50, 10, "Child care?", CC_check
  ButtonGroup ButtonPressed
    OkButton 175, 115, 50, 15
    CancelButton 230, 115, 50, 15
  Text 80, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  GroupBox 5, 5, 60, 120, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 95, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog

'Dialog asks what stats are being pulled
Do
	Dialog Dialog1
	cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

'CREATE ARRAY HERE

'Setting the variable for what's to come
all_case_numbers_array = "*"
case_count = 0

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

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "PND2")       'looking at PND2 to confirm day 30 AND look for MSA cases - which get 60 days
	EMWriteScreen worker, 21, 13
	transmit
	'This code is for bypassing a warning box if the basket has too many cases
	EMWaitReady 0, 0
	row = 1
	col = 1
	EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
	If row <> 0 THEN transmit
	'TODO add handling to read for an additional app line so that we are sure we are reading the correct line for days pending and cash program
	'Skips workers with no info
	EMReadScreen has_content_check, 6, 3, 74
	If has_content_check <> "0 Of 0" then
		'Grabbing each case number on screen
		Do
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5	'Reading case number
				EMReadScreen client_name, 22, MAXIS_row, 16		'Reading client name
				EMReadScreen APPL_date, 8, MAXIS_row, 38		'Reading application date
				EMReadScreen days_pending, 4, MAXIS_row, 49		'Reading days pending
				EMReadScreen cash_status, 1, MAXIS_row, 54		'Reading cash status
				EMReadScreen cash_prog, 2, MAXIS_row, 56
				EMReadScreen SNAP_status, 1, MAXIS_row, 62		'Reading SNAP status
				EMReadScreen HC_status, 1, MAXIS_row, 65		'Reading HC status
				EMReadScreen EA_status, 1, MAXIS_row, 68		'Reading EA status
				EMReadScreen GRH_status, 1, MAXIS_row, 72		'Reading GRH status
				EMReadScreen IVE_status, 1, MAXIS_row, 76		'Reading IV-E status
				EMReadScreen CC_status, 1, MAXIS_row, 80		'Reading CC status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				client_name = trim(client_name)
				MAXIS_case_number = trim(MAXIS_case_number)
				If client_name <> "ADDITIONAL APP" Then			'When there is an additional app on this rept, the script actually reads a case number even though one is not visible to the worker on the screen - so we are skipping this ghosting issue because it will ALWAYS find the previous case number.
					If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
					all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")
				End If

				If MAXIS_case_number = "" AND client_name = "" Then Exit Do			'Exits do if we reach the end

				'If additional application is rec'd then the excel output is the client's name, not ADDITIONAL APP
				if client_name = "ADDITIONAL APP" then
					EMReadScreen alt_client_name, 22, MAXIS_row - 1, 16
					client_name = "* " & trim(alt_client_name)                    'replaces alt name as the client name
				Else
					EMReadScreen next_client, 22, MAXIS_row + 1, 16
					next_client = trim(next_client)
					If next_client = "ADDITIONAL APP" Then client_name = "* " & client_name
				END IF

				'Cleaning up each program's status
				SNAP_status = trim(replace(SNAP_status, "_", ""))
				cash_status = trim(replace(cash_status, "_", ""))
				cash_prog = trim(replace(cash_prog, "_", ""))
				HC_status = trim(replace(HC_status, "_", ""))
				EA_status = trim(replace(EA_status, "_", ""))
				GRH_status = trim(replace(GRH_status, "_", ""))
				IVE_status = trim(replace(IVE_status, "_", ""))
				CC_status = trim(replace(CC_status, "_", ""))

				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
				If SNAP_status <> "" and SNAP_check = checked then add_case_info_to_Excel = True
				If cash_status <> "" and cash_check = checked then add_case_info_to_Excel = True
				If HC_status <> "" and HC_check = checked then add_case_info_to_Excel = True
				If EA_status <> "" and EA_check = checked then add_case_info_to_Excel = True
				If GRH_status <> "" and GRH_check = checked then add_case_info_to_Excel = True
				If IVE_status <> "" and IVE_check = checked then add_case_info_to_Excel = True
				If CC_status <> "" and CC_check = checked then add_case_info_to_Excel = True


				If add_case_info_to_Excel = True then
					ReDim preserve PND2_ARRAY(pnd2_last_const, case_count)

					PND2_ARRAY(worker_numb_const, case_count) 	= worker
					PND2_ARRAY(case_numb_const, case_count)		= MAXIS_case_number
					PND2_ARRAY(case_name_const, case_count)		= client_name
					PND2_ARRAY(appl_date_const, case_count)		= replace(APPL_date, " ", "/")
					PND2_ARRAY(days_pending_const, case_count)	= abs(days_pending)
					If SNAP_check = checked then PND2_ARRAY(snap_const, case_count)	= SNAP_status
					If cash_check = checked then
						PND2_ARRAY(cash_const, case_count)			= cash_status
						PND2_ARRAY(cash_prog_const, case_count)		= cash_prog
					End If
					If HC_check = checked then PND2_ARRAY(hc_const, case_count)		= HC_status
					If EA_check = checked then PND2_ARRAY(ea_const, case_count)		= EA_status
					If GRH_check = checked then PND2_ARRAY(grh_const, case_count)	= GRH_status
					If IVE_check = checked then PND2_ARRAY(ive_const, case_count)	= IVE_status
					If CC_check = checked then PND2_ARRAY(ccap_const, case_count)	= CC_status
					case_count = case_count + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next

'CREATE DIALOG WITH INFO
total_case_count = 0

snap_case_count = 0
snap_pending_0_20 = 0
snap_pending_21_29 = 0
snap_pending_30 = 0
snap_pending_31_45 = 0
snap_pending_over_45 = 0

cash_case_count = 0
cash_pending_0_20 = 0
cash_pending_21_29 = 0
cash_pending_30 = 0
cash_pending_31_45 = 0
cash_pending_over_45 = 0

hc_case_count = 0
hc_pending_0_20 = 0
hc_pending_21_29 = 0
hc_pending_30 = 0
hc_pending_31_45 = 0
hc_pending_over_45 = 0

emer_case_count = 0
emer_pending_0_20 = 0
emer_pending_21_29 = 0
emer_pending_30 = 0
emer_pending_31_45 = 0
emer_pending_over_45 = 0

grh_case_count = 0
grh_pending_0_20 = 0
grh_pending_21_29 = 0
grh_pending_30 = 0
grh_pending_31_45 = 0
grh_pending_over_45 = 0

ive_case_count = 0
ccap_case_count = 0

for each_case = 0 to UBound(PND2_ARRAY, 2)
	total_case_count = total_case_count + 1
	If SNAP_check = checked then
		If PND2_ARRAY(snap_const, each_case) <> "" Then
			snap_case_count = snap_case_count + 1
			If PND2_ARRAY(days_pending_const, each_case) < 21 Then snap_pending_0_20 = snap_pending_0_20 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 30 AND PND2_ARRAY(days_pending_const, each_case) > 20 Then snap_pending_21_29 = snap_pending_21_29 + 1
			If PND2_ARRAY(days_pending_const, each_case) = 30 Then snap_pending_30 = snap_pending_30 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 46 AND PND2_ARRAY(days_pending_const, each_case) > 30 Then snap_pending_31_45 = snap_pending_31_45 + 1
			If PND2_ARRAY(days_pending_const, each_case) > 45 Then snap_pending_over_45 = snap_pending_over_45 + 1
		End If
	End If

	If cash_check = checked then
		IF PND2_ARRAY(cash_const, each_case) <> "" Then
			cash_case_count = cash_case_count + 1
			If PND2_ARRAY(days_pending_const, each_case) < 21 Then cash_pending_0_20 = cash_pending_0_20 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 30 AND PND2_ARRAY(days_pending_const, each_case) > 20 Then cash_pending_21_29 = cash_pending_21_29 + 1
			If PND2_ARRAY(days_pending_const, each_case) = 30 Then cash_pending_30 = cash_pending_30 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 46 AND PND2_ARRAY(days_pending_const, each_case) > 30 Then cash_pending_31_45 = cash_pending_31_45 + 1
			If PND2_ARRAY(days_pending_const, each_case) > 45 Then cash_pending_over_45 = cash_pending_over_45 + 1
		End If
	End If

	If HC_check = checked then
		If PND2_ARRAY(hc_const, each_case)	<> "" Then
			hc_case_count = hc_case_count + 1
			If PND2_ARRAY(days_pending_const, each_case) < 21 Then hc_pending_0_20 = hc_pending_0_20 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 30 AND PND2_ARRAY(days_pending_const, each_case) > 20 Then hc_pending_21_29 = hc_pending_21_29 + 1
			If PND2_ARRAY(days_pending_const, each_case) = 30 Then hc_pending_30 = hc_pending_30 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 46 AND PND2_ARRAY(days_pending_const, each_case) > 30 Then hc_pending_31_45 = hc_pending_31_45 + 1
			If PND2_ARRAY(days_pending_const, each_case) > 45 Then hc_pending_over_45 = hc_pending_over_45 + 1
		End If
	End If

	If EA_check = checked then
		If PND2_ARRAY(ea_const, each_case) <> "" Then
			emer_case_count = emer_case_count + 1
			If PND2_ARRAY(days_pending_const, each_case) < 21 Then emer_pending_0_20 = emer_pending_0_20 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 30 AND PND2_ARRAY(days_pending_const, each_case) > 20 Then emer_pending_21_29 = emer_pending_21_29 + 1
			If PND2_ARRAY(days_pending_const, each_case) = 30 Then emer_pending_30 = emer_pending_30 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 46 AND PND2_ARRAY(days_pending_const, each_case) > 30 Then emer_pending_31_45 = emer_pending_31_45 + 1
			If PND2_ARRAY(days_pending_const, each_case) > 45 Then emer_pending_over_45 = emer_pending_over_45 + 1
		End If
	End If

	If GRH_check = checked then
		If PND2_ARRAY(grh_const, each_case) <> "" Then
			grh_case_count = grh_case_count + 1
			If PND2_ARRAY(days_pending_const, each_case) < 21 Then grh_pending_0_20 = grh_pending_0_20 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 30 AND PND2_ARRAY(days_pending_const, each_case) > 20 Then grh_pending_21_29 = grh_pending_21_29 + 1
			If PND2_ARRAY(days_pending_const, each_case) = 30 Then grh_pending_30 = grh_pending_30 + 1
			If PND2_ARRAY(days_pending_const, each_case) < 46 AND PND2_ARRAY(days_pending_const, each_case) > 30 Then grh_pending_31_45 = grh_pending_31_45 + 1
			If PND2_ARRAY(days_pending_const, each_case) > 45 Then grh_pending_over_45 = grh_pending_over_45 + 1
		End If
	End If

	If IVE_check = checked then
		If PND2_ARRAY(ive_const, each_case) <> "" Then
			ive_case_count = ive_case_count + 1
		End If
	End If

	If CC_check = checked then
		If PND2_ARRAY(ccap_const, each_case) <> "" Then
			ccap_case_count = ccap_case_count + 1
		End If
	End If

Next

dlg_len = 20
If SNAP_check = checked then
	dlg_len = dlg_len + 50
	If snap_case_count <> 0 Then
		snap_pending_0_20_percent = (snap_pending_0_20 / snap_case_count)*100
		snap_pending_0_20_percent = FormatNumber(snap_pending_0_20_percent, 2, -1, 0, -1)
		snap_pending_21_29_percent = (snap_pending_21_29 / snap_case_count)*100
		snap_pending_21_29_percent = FormatNumber(snap_pending_21_29_percent, 2, -1, 0, -1)
		snap_pending_30_percent = (snap_pending_30 / snap_case_count)*100
		snap_pending_30_percent = FormatNumber(snap_pending_30_percent, 2, -1, 0, -1)
		snap_pending_31_45_percent = (snap_pending_31_45 / snap_case_count)*100
		snap_pending_31_45_percent = FormatNumber(snap_pending_31_45_percent, 2, -1, 0, -1)
		snap_pending_over_45_percent = (snap_pending_over_45 / snap_case_count)*100
		snap_pending_over_45_percent = FormatNumber(snap_pending_over_45_percent, 2, -1, 0, -1)
	End If
End If
If cash_check = checked then
	dlg_len = dlg_len + 50
	If cash_case_count <> 0 Then
		cash_pending_0_20_percent = (cash_pending_0_20 / cash_case_count)*100
		cash_pending_0_20_percent = FormatNumber(cash_pending_0_20_percent, 2, -1, 0, -1)
		cash_pending_21_29_percent = (cash_pending_21_29 / cash_case_count)*100
		cash_pending_21_29_percent = FormatNumber(cash_pending_21_29_percent, 2, -1, 0, -1)
		cash_pending_30_percent = (cash_pending_30 / cash_case_count)*100
		cash_pending_30_percent = FormatNumber(cash_pending_30_percent, 2, -1, 0, -1)
		cash_pending_31_45_percent = (cash_pending_31_45 / cash_case_count)*100
		cash_pending_31_45_percent = FormatNumber(cash_pending_31_45_percent, 2, -1, 0, -1)
		cash_pending_over_45_percent = (cash_pending_over_45 / cash_case_count)*100
		cash_pending_over_45_percent = FormatNumber(cash_pending_over_45_percent, 2, -1, 0, -1)
	End If
End If
If HC_check = checked then
	dlg_len = dlg_len + 50
	If hc_case_count <> 0 Then
		hc_pending_0_20_percent = (hc_pending_0_20 / hc_case_count)*100
		hc_pending_0_20_percent = FormatNumber(hc_pending_0_20_percent, 2, -1, 0, -1)
		hc_pending_21_29_percent = (hc_pending_21_29 / hc_case_count)*100
		hc_pending_21_29_percent = FormatNumber(hc_pending_21_29_percent, 2, -1, 0, -1)
		hc_pending_30_percent = (hc_pending_30 / hc_case_count)*100
		hc_pending_30_percent = FormatNumber(hc_pending_30_percent, 2, -1, 0, -1)
		hc_pending_31_45_percent = (hc_pending_31_45 / hc_case_count)*100
		hc_pending_31_45_percent = FormatNumber(hc_pending_31_45_percent, 2, -1, 0, -1)
		hc_pending_over_45_percent = (hc_pending_over_45 / hc_case_count)*100
		hc_pending_over_45_percent = FormatNumber(hc_pending_over_45_percent, 2, -1, 0, -1)
	End If
End If
If EA_check = checked then
	dlg_len = dlg_len + 50
	If emer_case_count <> 0 Then
		emer_pending_0_20_percent = (emer_pending_0_20 / emer_case_count)*100
		emer_pending_0_20_percent = FormatNumber(emer_pending_0_20_percent, 2, -1, 0, -1)
		emer_pending_21_29_percent = (emer_pending_21_29 / emer_case_count)*100
		emer_pending_21_29_percent = FormatNumber(emer_pending_21_29_percent, 2, -1, 0, -1)
		emer_pending_30_percent = (emer_pending_30 / emer_case_count)*100
		emer_pending_30_percent = FormatNumber(emer_pending_30_percent, 2, -1, 0, -1)
		emer_pending_31_45_percent = (emer_pending_31_45 / emer_case_count)*100
		emer_pending_31_45_percent = FormatNumber(emer_pending_31_45_percent, 2, -1, 0, -1)
		emer_pending_over_45_percent = (emer_pending_over_45 / emer_case_count)*100
		emer_pending_over_45_percent = FormatNumber(emer_pending_over_45_percent, 2, -1, 0, -1)
	End If
End If
If GRH_check = checked then
	dlg_len = dlg_len + 50
	If grh_case_count <> 0 Then
		grh_pending_0_20_percent = (grh_pending_0_20 / grh_case_count)*100
		grh_pending_0_20_percent = FormatNumber(grh_pending_0_20_percent, 2, -1, 0, -1)
		grh_pending_21_29_percent = (grh_pending_21_29 / grh_case_count)*100
		grh_pending_21_29_percent = FormatNumber(grh_pending_21_29_percent, 2, -1, 0, -1)
		grh_pending_30_percent = (grh_pending_30 / grh_case_count)*100
		grh_pending_30_percent = FormatNumber(grh_pending_30_percent, 2, -1, 0, -1)
		grh_pending_31_45_percent = (grh_pending_31_45 / grh_case_count)*100
		grh_pending_31_45_percent = FormatNumber(grh_pending_31_45_percent, 2, -1, 0, -1)
		grh_pending_over_45_percent = (grh_pending_over_45 / grh_case_count)*100
		grh_pending_over_45_percent = FormatNumber(grh_pending_over_45_percent, 2, -1, 0, -1)
	End If
End If
If IVE_check = checked or CC_check = checked then
	dlg_len = dlg_len + 20
End If


If IVE_check = checked then
End If
If CC_check = checked then
End If

y_pos = 25
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, dlg_len, "CASE COUNT"
ButtonGroup ButtonPressed
  OkButton 310, 5, 50, 15
  Text 10, 10, 110, 10, "All pending cases: " & total_case_count
	If SNAP_check = checked then
		GroupBox 10, y_pos, 350, 45, "SNAP"
		y_pos = y_pos + 10
		Text 20, y_pos, 	75, 10, "Pending SNAP Cases:"
		Text 20, y_pos+15,	75, 10, snap_case_count
		Text 95, y_pos, 	40, 10, "0 - 20 Days"
		Text 95, y_pos+15,	50, 10, snap_pending_0_20 & " - " & snap_pending_0_20_percent & " %"
		Text 145, y_pos, 	40, 10, "21-29 Days"
		Text 145, y_pos+15,	50, 10, snap_pending_21_29 & " - " & snap_pending_21_29_percent & " %"
		Text 195, y_pos, 	35, 10, "30 Days"
		Text 195, y_pos+15,	50, 10, snap_pending_30 & " - " & snap_pending_30_percent & " %"
		Text 245, y_pos, 	45, 10, "31 - 45 Days"
		Text 245, y_pos+15,	50, 10, snap_pending_31_45 & " - " & snap_pending_31_45_percent & " %"
		Text 295, y_pos, 	55, 10, "45 or More Days"
		Text 295, y_pos+15,	45, 10, snap_pending_over_45 & " - " & snap_pending_over_45_percent & " %"
		y_pos = y_pos + 40
	End If

	If cash_check = checked then
		GroupBox 10, y_pos, 350, 45, "CASH"
		y_pos = y_pos + 10
		Text 20, y_pos, 	75, 10, "Pending Cash Cases: "
		Text 20, y_pos+15,	75, 10, cash_case_count
		Text 95, y_pos, 	40, 10, "0 - 20 Days"
		Text 95, y_pos+15,	50, 10, cash_pending_0_20 & " - " & cash_pending_0_20_percent & " %"
		Text 145, y_pos, 	40, 10, "21-29 Days"
		Text 145, y_pos+15,	50, 10, cash_pending_21_29 & " - " & cash_pending_21_29_percent & " %"
		Text 195, y_pos, 	35, 10, "30 Days"
		Text 195, y_pos+15,	50, 10, cash_pending_30 & " - " & cash_pending_30_percent & " %"
		Text 245, y_pos, 	45, 10, "31 - 45 Days"
		Text 245, y_pos+15,	50, 10, cash_pending_31_45 & " - " & cash_pending_31_45_percent & " %"
		Text 295, y_pos, 	55, 10, "45 or More Days"
		Text 295, y_pos+15,	45, 10, cash_pending_over_45 & " - " & cash_pending_over_45_percent & " %"
		y_pos = y_pos + 40
	End If

	If HC_check = checked then
		GroupBox 10, y_pos, 350, 45, "HC"
		y_pos = y_pos + 10
		Text 20, y_pos, 	75, 10, "Pending HC Cases: "
		Text 20, y_pos+15,	75, 10, hc_case_count
		Text 95, y_pos, 	40, 10, "0 - 20 Days"
		Text 95, y_pos+15,	50, 10, hc_pending_0_20 & " - " & hc_pending_0_20_percent & " %"
		Text 145, y_pos, 	40, 10, "21-29 Days"
		Text 145, y_pos+15,	50, 10, hc_pending_21_29 & " - " & hc_pending_21_29_percent & " %"
		Text 195, y_pos, 	35, 10, "30 Days"
		Text 195, y_pos+15,	50, 10, hc_pending_30 & " - " & hc_pending_30_percent & " %"
		Text 245, y_pos, 	45, 10, "31 - 45 Days"
		Text 245, y_pos+15,	50, 10, hc_pending_31_45 & " - " & hc_pending_31_45_percent & " %"
		Text 295, y_pos, 	55, 10, "45 or More Days"
		Text 295, y_pos+15,	45, 10, hc_pending_over_45 & " - " & hc_pending_over_45_percent & " %"
		y_pos = y_pos + 40
	End If

	If EA_check = checked then
		GroupBox 10, y_pos, 350, 45, "EMER"
		y_pos = y_pos + 10
		Text 20, y_pos, 	75, 10, "Pending EMER Cases: "
		Text 20, y_pos+15,	75, 10, emer_case_count
		Text 95, y_pos, 	40, 10, "0 - 20 Days"
		Text 95, y_pos+15,	50, 10, emer_pending_0_20 & " - " & emer_pending_0_20_percent & " %"
		Text 145, y_pos, 	40, 10, "21-29 Days"
		Text 145, y_pos+15,	50, 10, emer_pending_21_29 & " - " & emer_pending_21_29_percent & " %"
		Text 195, y_pos, 	35, 10, "30 Days"
		Text 195, y_pos+15,	50, 10, emer_pending_30 & " - " & emer_pending_30_percent & " %"
		Text 245, y_pos, 	45, 10, "31 - 45 Days"
		Text 245, y_pos+15,	50, 10, emer_pending_31_45 & " - " & emer_pending_31_45_percent & " %"
		Text 295, y_pos, 	55, 10, "45 or More Days"
		Text 295, y_pos+15,	45, 10, emer_pending_over_45 & " - " & emer_pending_over_45_percent & " %"
		y_pos = y_pos + 40
	End If

	If GRH_check = checked then
		GroupBox 10, y_pos, 350, 45, "GRH"
		y_pos = y_pos + 10
		Text 20, y_pos, 	75, 10, "Pending GRH Cases: "
		Text 20, y_pos+15,	75, 10, grh_case_count
		Text 95, y_pos, 	40, 10, "0 - 20 Days"
		Text 95, y_pos+15,	50, 10, grh_pending_0_20 & " - " & grh_pending_0_20_percent & " %"
		Text 145, y_pos, 	40, 10, "21-29 Days"
		Text 145, y_pos+15,	50, 10, grh_pending_21_29 & " - " & grh_pending_21_29_percent & " %"
		Text 195, y_pos, 	35, 10, "30 Days"
		Text 195, y_pos+15,	50, 10, grh_pending_30 & " - " & grh_pending_30_percent & " %"
		Text 245, y_pos, 	45, 10, "31 - 45 Days"
		Text 245, y_pos+15,	50, 10, grh_pending_31_45 & " - " & grh_pending_31_45_percent & " %"
		Text 295, y_pos, 	55, 10, "45 or More Days"
		Text 295, y_pos+15,	45, 10, grh_pending_over_45 & " - " & grh_pending_over_45_percent & " %"
		y_pos = y_pos + 40
	End If
	If IVE_check = checked then
		Text 15, y_pos, 95, 10, "Pending IV-E Cases: " & ive_case_count
	End If
	If CC_check = checked then
		Text 130, y_pos, 95, 10, "Pending CCAP Cases: " & ccap_case_count
	End If
EndDialog

dialog Dialog1




call script_end_procedure("")