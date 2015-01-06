'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    HOW THE DAIL SCRUBER WORKS:
'
'    This script opens up other script files, using a custom function (run_DAIL_scrubber_script), followed by the path to the script file. It's done this
'      way because there could be hundreds of DAIL messages, and to work all of the combinations into one script would be incredibly tedious and long.
'
'    This script works by moving the message (where the cursor is located) to the top of the screen, and then reading the message text. Whatever the
'      message text says dictates which script loads up.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - DAIL SCRUBBER.vbs"
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

'CONNECTS TO DEFAULT SCREEN
EMConnect ""

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
transmit

'THE FOLLOWING CODES ARE THE INDIVIDUAL MESSAGES. IT READS THE MESSAGE, THEN CALLS A NEW SCRIPT.----------------------------------------------------------------------------------------------------

'Random messages generated from an affiliated case (loads AFFILIATED CASE LOOKUP)
EMReadScreen stat_check, 4, 6, 6
If stat_check = "FS  " or stat_check = "HC  " or stat_check = "GA  " or stat_check = "MSA " or stat_check = "STAT" then call run_from_GitHub(script_repository & "DAIL/DAIL - AFFILIATED CASE LOOKUP.vbs")

'RSDI/BENDEX info received by agency (loads BNDX SCRUBBER)
EMReadScreen BENDEX_check, 47, 6, 30
If BENDEX_check = "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC" then call run_from_GitHub(script_repository & "DAIL/DAIL - BNDX SCRUBBER.vbs")

'CIT/ID has been verified through the SSA (loads CITIZENSHIP VERIFIED)
EMReadScreen CIT_check, 46, 6, 20
If CIT_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then call run_from_GitHub(script_repository & "DAIL/DAIL - CITIZENSHIP VERIFIED.vbs")

'CS reports a new employer to the worker (loads CS REPORTED NEW EMPLOYER)
EMReadScreen CS_new_emp_check, 25, 6, 20
If CS_new_emp_check = "CS REPORTED: NEW EMPLOYER" then call run_from_GitHub(script_repository & "DAIL/DAIL - CS REPORTED NEW EMPLOYER.vbs")

'Child support messages (loads CSES PROCESSING)
EMReadScreen CSES_check, 4, 6, 6
If CSES_check = "CSES" then
  EMReadScreen CSES_DISB_check, 4, 6, 20
  If CSES_DISB_check = "DISB" then call run_from_GitHub(script_repository & "DAIL/DAIL - CSES PROCESSING.vbs")
End if

'Disability certification ends in 60 days (loads DISA MESSAGE)
EMReadScreen DISA_check, 58, 6, 20
If DISA_check = "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" then call run_from_GitHub(script_repository & "DAIL/DAIL - DISA MESSAGE.vbs")

'Client can receive an FMED deduction for SNAP (loads FMED DEDUCTION)
EMReadScreen FMED_check, 59, 6, 20
If FMED_check = "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" then call run_from_GitHub(script_repository & "DAIL/DAIL - FMED DEDUCTION.vbs")

'Remedial care messages. May only happen at COLA (loads LTC - REMEDIAL CARE)
EMReadScreen remedial_care_check, 41, 6, 20
If remedial_care_check = "REF 01 PERSON HAS REMEDIAL CARE DEDUCTION" then call run_from_GitHub(script_repository & "DAIL/DAIL - LTC - REMEDIAL CARE.vbs")

'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 15, 6, 20
If HIRE_check = "NEW JOB DETAILS" then call run_from_GitHub(script_repository & "DAIL/DAIL - NEW HIRE.vbs")

'SSI info received by agency (loads SDX INFO HAS BEEN STORED)
EMReadScreen SDX_check, 44, 6, 30
If SDX_check = "SDX INFORMATION HAS BEEN STORED - CHECK INFC" then call run_from_GitHub(script_repository & "DAIL/DAIL - SDX INFO HAS BEEN STORED.vbs")

'Student income is ending (loads STUDENT INCOME)
EMReadScreen SCHL_check, 58, 6, 20
If SCHL_check = "STUDENT INCOME HAS ENDED - REVIEW FS AND/OR HC RESULTS/APP" then call run_from_GitHub(script_repository & "DAIL/DAIL - STUDENT INCOME.vbs")

'SSA info received by agency (loads TPQY RESPONSE)
EMReadScreen TPQY_check, 31, 6, 30
If TPQY_check = "TPQY RESPONSE RECEIVED FROM SSA" then call run_from_GitHub(script_repository & "DAIL/DAIL - TPQY RESPONSE.vbs")

'NOW IF NO SCRIPT HAS BEEN WRITTEN FOR IT, THE DAIL SCRUBBER STOPS AND GENERATES A MESSAGE TO THE WORKER.----------------------------------------------------------------------------------------------------
script_end_procedure("You are not on a supported DAIL message. The script will now stop.")
