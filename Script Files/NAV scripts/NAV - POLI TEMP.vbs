'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - POLI TEMP"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS--------------------------------------------------
BeginDialog POLI_TEMP_dialog, 0, 0, 256, 60, "POLI/TEMP dialog"
  OptionGroup RadioGroup1
    RadioButton 5, 30, 175, 10, "Table of Contents (search by TEMP section code)", table_radio
    RadioButton 5, 45, 150, 10, "Index of Topics (search by a word or topic)", index_radio
  ButtonGroup ButtonPressed
    OkButton 195, 10, 50, 15
    CancelButton 195, 30, 50, 15
  Text 10, 10, 160, 10, "What area of POLI/TEMP do you want to go to?"
EndDialog


'THE SCRIPT

'Displays dialog
Dialog POLI_TEMP_dialog
If buttonpressed = cancel then stopscript

'Determines which POLI/TEMP section to go to, using the radioboxes outcome to decide
If radiogroup1 = table_radio then 
	panel_title = "TABLE"
ElseIf radiogroup1 = index_radio then
	panel_title = "INDEX"
End if


'Connects to BlueZone
EMConnect ""

'Checks to make sure we're in MAXIS
MAXIS_check_function

'Navigates to POLI (can't direct navigate to TEMP)
call navigate_to_screen("POLI", "____")

'Writes TEMP
EMWriteScreen "TEMP", 5, 40

'Writes the panel_title selection
EMWriteScreen panel_title, 21, 71

'Transmits
transmit