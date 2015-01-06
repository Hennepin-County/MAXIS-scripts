'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - POLI-TEMP.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

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
