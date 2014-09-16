'Informational front-end message, date dependent.
If datediff("d", "04/02/2013", now) < 5 then MsgBox "This script has been updated as of 04/02/2013! There's now a checkbox for starting the denied programs script right from this one."

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "SCRIPT NAME HERE"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

Set MNSURE_FUNCTIONS_fso = CreateObject("Scripting.FileSystemObject")
Set fso_MNSURE_FUNCTIONS_command = MNSURE_FUNCTIONS_fso.OpenTextFile("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\MNSURE FUNCTIONS FILE.vbs")
MNSURE_FUNCTIONS_contents = fso_MNSURE_FUNCTIONS_command.ReadAll
fso_MNSURE_FUNCTIONS_command.Close
Execute MNSURE_FUNCTIONS_contents 

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog notice_generator_selector, 0, 0, 150, 76, "MNSure Notice Generator"
  ButtonGroup ButtonPressed
    PushButton 4, 26, 69, 29, "Single Notice", run_single_notice
    PushButton 78, 26, 69, 29, "Mass Notice's", run_mass_notices
    PushButton 108, 60, 39, 12, "Stopscript", stopscript_button
  Text 8, 5, 139, 16, "Please select the type of notice generator you would like to use."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

Dialog notice_generator_selector
	If buttonpressed = stopscript_button then
		stopscript
	ElseIf buttonpressed = run_single_notice then
		call run_file("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\NOTICES - Single Notice Generator.vbs")
	ElseIf buttonpressed = run_mass_notices then
		call run_file("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\NOTICES - Mass Notice Generator.vbs")
End If

script_end_procedure("")