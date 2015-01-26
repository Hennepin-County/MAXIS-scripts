'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CHECK EDRS.vbs"
start_time = timer

''LOADING ROUTINE FUNCTIONS
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0") 'Creates an object to get a URL
req.open "GET", url, FALSE	'Attempts to open the URL
req.send 'Sends request
IF req.Status = 200 THEN	'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject") 'Creates an FSO
	Execute req.responseText 'Executes the script code
ELSE	'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
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


BeginDialog EDRS_dialog, 0, 0, 156, 80, "EDRS dialog"
  EditBox 60, 10, 80, 15, case_number
  EditBox 60, 30, 25, 15, memb_number
  ButtonGroup ButtonPressed
    OkButton 15, 55, 50, 15
    CancelButton 80, 55, 50, 15
  Text 5, 15, 50, 10, "Case Number:"
  Text 5, 35, 50, 10, "Memb Number:"
EndDialog





EMConnect ""

'Hunts for Maxis case number to autofill it
Call MAXIS_case_number_finder(case_number)


DO
	dialog EDRS_dialog
	IF buttonpressed = 0 THEN stopscript
	IF case_number = "" THEN MSGBOX "Please enter a case number"
	IF memb_number = "" THEN MSGBOX "Please enter a member number"

LOOP UNTIL case_number <> "" AND memb_number <> ""

'changing footer dates to current month to avoid invalid months. 
footer_month = datepart("M", date)
	IF Len(footer_month) <> 2 THEN footer_month = "0" & footer_month 
footer_year = right(datepart("YYYY", date), 2)

'error proofs for 1 digit member numbers
IF LEN(memb_number) <> 2 THEN memb_number = "0" & memb_number

'Error proof functions
Maxis_check_function
MAXIS_background_check



'Navigate to stat/memb and check for ERRR message
CALL Navigate_to_screen("STAT", "MEMB")
ERRR_screen_check

'Navigating to selected memb panel
EMwritescreen memb_number, 20, 76
transmit

EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
IF no_MEMB = "Arrival Date:" THEN script_end_procedure("This HH member does not exist.")

'Reading SSN number and removing spaces
Emreadscreen SSN_number, 11, 7, 42  
SSN_number = replace(SSN_number, " ", "")
'Reading Last name and removing spaces
EMReadscreen Last_name, 25, 6, 30
Last_name = replace(Last_name, "_", "")
'Reading First name and removing spaces
EMReadscreen First_name, 12, 6, 63
First_name = replace(First_name, "_", "")
'Reading Middle initial and replacing _ with a blank if empty. 
EMReadscreen Middle_initial, 1, 6, 79
Middle_initial = replace(Middle_initial, "_", "")


'Navigate back to self and to EDRS
Back_to_self
CALL Navigate_to_screen("INFC", "EDRS")

'Write in SSN number into EDRS
EMwritescreen SSN_number, 2, 7
transmit
Emreadscreen SSN_output, 7, 24, 2

'Check to see what results you get from entering the SSN. If you get NO DISQ then check the person's name
IF SSN_output = "NO DISQ" THEN
	EMWritescreen Last_name, 2, 24
	EMWritescreen First_name, 2, 58
	EMWritescreen Middle_initial, 2, 76
	transmit
	EMreadscreen NAME_output, 7, 24, 2
	IF NAME_output = "NO DISQ" THEN        'If after entering a name you still get NO DISQ then let worker know otherwise let them know you found a name. 
		MSGBOX "No disqualifications found for " & First_name & " " & Last_name & " Member #: " & Memb_number
	ELSE
		MSGBOX "Client's name has a match"
	END IF
ELSE
	MSGBOX "SSN number has a match"        'If after searching a SSN number you don't get the NO DISQ message then let worker know you found the SSN
END IF


script_end_procedure("")

