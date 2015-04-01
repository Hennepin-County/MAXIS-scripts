'STATS GATHERING ----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - LTC - TRANSFER PENALTY.vbs"
start_time = timer
'Reference source: http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=dhs16_150210	

'DIM beta_agency

'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
'DIM url, req, fso
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
SET req = CreateObject("Msxml2.XMLHttp.6.0")					'Creates an object to get a URL
req.open "GET", url, FALSE										'Attempts to open the URL
req.send														'Sends request
IF req.Status = 200 THEN										'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")		'Creates an FSO
	Execute req.responseText									'Executes the script code
ELSE															'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
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

'DIM name_of_script
'DIM start_time
'DIM case_number
'DIM ButtonPressed
'DIM case_number_dialogbox
'DIM LTC_transfer_penalty_dialog
'DIM type_of_transfer_list
'DIM transfer_date
'DIM transfer_amount
'DIM date_of_application
'DIM baseline_date
'DIM date_client_was_otherwise_eligible
'DIM period_begins
'DIM last_full_month_of_period
'DIM partial_penalty_amount
'DIM other_information
'DIM harship_waiver_requested_check
'DIM hardship_waiver_approved_check
'DIM harship_waiver_details
'DIM case_action
'DIM worker_signature
'DIM lookback_period
'DIM end_of_lookback
'DIM row_1
'DIM row_3
'DIM row_4
'DIM row_5
'DIM row_6
'DIM row_8
'DIM row_9

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialogbox, 0, 0, 191, 80, "Dialog"		'dialog box where worker enters the case number (and at some point applicable month & year)
  EditBox 75, 15, 80, 15, case_number					'once worker selects ok, it will move to the next dialog box.  If worker selects cancel, then
  'EditBox 75, 35, 40, 15, month_editbox					'script will end
  'EditBox 120, 35, 35, 15, year_editbox
  ButtonGroup ButtonPressed
    OkButton 80, 60, 50, 15
    CancelButton 135, 60, 50, 15
  'Text 10, 35, 60, 15, "Footer month:"
  Text 10, 15, 60, 15, "Case number: "
EndDialog

BeginDialog LTC_transfer_penalty_dialog, 0, 0, 226, 375, "Dialog"
  DropListBox 65, 5, 125, 15, "Annuity"+chr(9)+"Life Estate"+chr(9)+"Uncompensated Transfer"+chr(9)+"Other", type_of_transfer_list
  EditBox 55, 25, 45, 15, transfer_date
  EditBox 165, 25, 45, 15, transfer_amount
  EditBox 165, 45, 45, 15, date_of_application
  EditBox 165, 65, 45, 15, baseline_date
  EditBox 165, 85, 45, 15, date_client_was_otherwise_eligible
  EditBox 165, 105, 45, 15, period_begins
  EditBox 165, 125, 45, 15, last_full_month_of_period
  EditBox 165, 145, 45, 15, partial_penalty_amount
  EditBox 70, 165, 140, 15, other_information
  CheckBox 5, 185, 100, 10, "Hardship waiver requested", harship_waiver_requested_check
  CheckBox 115, 185, 100, 10, "Hardship waiver approved", hardship_waiver_approved_check
  EditBox 85, 200, 125, 15, harship_waiver_details
  EditBox 50, 220, 160, 15, case_action
  EditBox 65, 240, 45, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 115, 240, 50, 15
    CancelButton 170, 240, 50, 15
  Text 5, 5, 55, 15, "Type of transfer: "
  Text 5, 25, 45, 15, "Transfer date:"
  Text 105, 25, 55, 15, "Transfer amount:"
  Text 5, 45, 115, 15, "Date of application or LTC request:"
  Text 5, 65, 105, 15, "Baseline date"
  Text 5, 85, 120, 15, "Date client was otherwise eligible:"
  Text 5, 105, 80, 15, "Transfer period begins:"
  Text 5, 125, 110, 15, "Last full month of transfer period:"
  Text 5, 145, 80, 15, "Partial penalty amount:"
  Text 5, 165, 60, 15, "Other information:"
  Text 5, 200, 80, 15, "Hardship waiver details:"
  Text 5, 220, 40, 15, "Case action:"
  Text 5, 240, 60, 15, "Worker signature:"
  Text 10, 295, 200, 40, "1. A person is residing in an LTCF or, for a person requesting services through a home and community-based waiver program, the date a screening occurred that indicated a need for services provided through a home and community-based services waiver program AND"
  Text 10, 270, 195, 15, "Per HCPM 19.40.15: The baseline date is the date in which both of the following conditions are met: "
  Text 10, 340, 200, 20, "2. The person's initial request month for MA payment of LTC services"
  GroupBox 0, 260, 220, 105, ""
EndDialog




'SCRIPT BODY----------------------------------------------------------------------------------------------------
EMConnect ""		'Connecting to Bluezone

call MAXIS_case_number_finder(case_number)							'function autofills case number that worker already has on MAXIS screen

DO
		Dialog case_number_dialogbox									'calls up dialog for worker to enter case number and applicable month and year.  Script will 'loop' 
		IF buttonpressed = cancel THEN StopScript						'and verbally request the worker to enter a case number until the worker enters a case number.
		IF case_number = "" THEN MsgBox "You must enter a case number"
	LOOP UNTIL case_number <> ""
	
Call check_for_MAXIS(true)											'ensures that worker has not "passworded" out of MAXIS

DO
	DO
		Dialog LTC_transfer_penalty_dialog
		IF len(baseline_date) < 6 THEN MsgBox "You must enter a date in format MM/DD/YYYY"
	LOOP until len(baseline_date) >= 6
	IF worker_signature = "" THEN MsgBox "You must sign your case note!"
LOOP UNTIL worker_signature <> ""

Call navigate_to_screen ("elig", "HC__")
EMWriteScreen "x", 8, 26 
transmit

EMWriteScreen "x", 7, 17 
transmit

EMWriteScreen "x", 18, 3 
transmit

EMReadScreen row_1, 71, 5, 6
EMReadScreen row_3, 71, 7, 6
EMReadScreen row_4, 71, 8, 6
EMReadScreen row_5, 71, 9, 6
EMReadScreen row_6, 71, 10, 6 
EMReadScreen row_8, 71, 12, 6
EMReadScreen row_9, 71, 13, 6 


Call navigate_to_screen ("case", "note")						'function to navigate user to case note
PF9																	'brings case note into edit mode

'Autofill for the application_date variable, then determines lookback period based on the info
If baseline_date <> "" then lookback_period = dateadd("m", -60, cdate(baseline_date)) & ""

'Lookback period end date
If baseline_date <> "" then end_of_lookback = dateadd ("d", -1, cdate (baseline_date))

'Dollar bill symbol will be added to numeric variables 
IF transfer_amount <> "" THEN transfer_amount = "$" & transfer_amount
IF partial_penalty_amount <> "" THEN partial_penalty_amount = "$" & partial_penalty_amount

Call write_variable_in_case_note ("~~~TRANSFER PENALTY~~~")     			'adding information to case note
Call write_bullet_and_variable_in_case_note ("Type of transfer", type_of_transfer_list )        
Call write_bullet_and_variable_in_case_note ("Transfer date", transfer_date)         
Call write_bullet_and_variable_in_case_note ("Transfer amount", transfer_amount)          
Call write_bullet_and_variable_in_case_note ("Date of application or LTC request", date_of_application)                
Call write_bullet_and_variable_in_case_note ("Baseline Date", baseline_date)                   
Call write_bullet_and_variable_in_case_note ("Date client was otherwise eligible", date_client_was_otherwise_eligible) 
Call write_bullet_and_variable_in_case_note ("Lookback period", lookback_period & "-" & end_of_lookback)
Call write_bullet_and_variable_in_case_note ("Transfer period begins", period_begins) 
Call write_bullet_and_variable_in_case_note ("Last full month of transfer period", last_full_month_of_period)
Call write_bullet_and_variable_in_case_note ("Partial penalty amount", partial_penalty_amount)
Call write_bullet_and_variable_in_case_note ("Other information", other_information)                
IF harship_waiver_requested_check = 1 THEN Call write_variable_in_case_note ("* Hardship waiver requested")             
IF hardship_waiver_approved_check = 1 THEN Call write_variable_in_case_note ("* Hardship waiver approved")
Call write_bullet_and_variable_in_case_note ("Hardship waiver details", harship_waiver_details) 
Call write_bullet_and_variable_in_case_note ("Case Action", case_action) 
Call write_variable_in_case_note ("---") 
Call write_variable_in_case_note (row_1)
Call write_variable_in_case_note (row_3)
Call write_variable_in_case_note (row_4)
Call write_variable_in_case_note (row_5)
Call write_variable_in_case_note (row_6)
Call write_variable_in_case_note (row_8)
Call write_variable_in_case_note (row_9)
Call write_variable_in_case_note ("---")                         
call write_variable_in_case_note (worker_signature)

script_end_procedure ("")							'closing script and writing stats