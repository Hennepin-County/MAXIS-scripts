'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - pull data into Excel"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CUSTOM FUNCTION (WILL INCLUDE IN FUNCTIONS FILE BEFORE FULL DEPLOYMENT)



'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog rept_scanning_dialog, 0, 0, 136, 75, "REPT scanning dialog"
  ButtonGroup ButtonPressed
    PushButton 35, 20, 25, 10, "ACTV", ACTV_button
    PushButton 35, 30, 25, 10, "ARST", ARST_button
    PushButton 35, 40, 25, 10, "EOMC", EOMC_button
    PushButton 70, 20, 25, 10, "PND2", PND2_button
    PushButton 70, 30, 25, 10, "REVS", REVS_button
    PushButton 70, 40, 25, 10, "REVW", REVW_button
    CancelButton 40, 55, 50, 15
  Text 5, 5, 125, 10, "What area of REPT are you scanning?"
EndDialog

BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 286, 115, "Pull REPT data into Excel dialog"
  EditBox 150, 20, 130, 15, worker_number
  CheckBox 70, 55, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 35, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 50, 40, 10, "Cash?", cash_check
  CheckBox 10, 65, 40, 10, "HC?", HC_check
  CheckBox 10, 80, 40, 10, "EA?", EA_check
  CheckBox 10, 95, 40, 10, "GRH?", GRH_check
  ButtonGroup ButtonPressed
    OkButton 175, 95, 50, 15
    CancelButton 230, 95, 50, 15
  GroupBox 5, 20, 60, 90, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 70, 215, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 80, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 40, 215, 10, "Enter workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog



'VARIABLES TO DECLARE
all_case_numbers_array = " "					'Creating blank variable for the future array
call worker_county_code_determination(worker_county_code, two_digit_county_code)	'Determines worker county code
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows report scanning dialog, which asks user which report to generate.
dialog rept_scanning_dialog
If buttonpressed = cancel then stopscript

'Connecting to BlueZone
EMConnect ""

If buttonpressed = ACTV_button then call run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\BULK - REPT-ACTV list.vbs")
If buttonpressed = ARST_button then call run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\BULK - REPT-ARST list.vbs")
If buttonpressed = EOMC_button then call run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\BULK - REPT-EOMC list.vbs")
If buttonpressed = PND2_button then call run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\BULK - REPT-PND2 list.vbs")
If buttonpressed = REVS_button then call run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\BULK - REPT-REVS list.vbs")
If buttonpressed = REVW_button then call run_another_script("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\BULK - REPT-REVW list.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")