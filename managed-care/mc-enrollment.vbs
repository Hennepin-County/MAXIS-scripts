'**********THIS IS A RAMSEY SPECIFIC SCRIPT.  IF YOU REVERSE ENGINEER THIS SCRIPT, JUST BE CAREFUL.************
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Action - Managed Care Enrollment.vbs"
start_time = timer

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog case_dlg, 0, 0, 161, 150, "Enrollment Information"
  EditBox 90, 25, 60, 15, MMIS_case_number
  EditBox 90, 45, 25, 15, enrollment_month
  EditBox 115, 45, 25, 15, enrollment_year
  DropListBox 55, 75, 95, 15, "Select one..."+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Hennepin Health SNBC"+chr(9)+"Ucare", Health_plan
  CheckBox 120, 95, 25, 10, "Yes", Insurance_yes
  CheckBox 120, 105, 25, 10, "Yes", foster_care_yes
  ButtonGroup ButtonPressed
    OkButton 45, 125, 50, 15
    CancelButton 100, 125, 50, 15
  GroupBox 5, 10, 150, 55, "Leading zeros not needed"
  Text 10, 30, 50, 10, "Case Number:"
  Text 10, 50, 80, 10, "Enrollment Month/Year:"
  Text 10, 80, 40, 10, "Health plan:"
  Text 10, 95, 100, 10, "Other Insurance for this case?"
  Text 10, 105, 50, 10, "Foster Care?"
  ButtonGroup ButtonPressed
    OkButton 155, 330, 50, 15
EndDialog

BeginDialog RPPH_error_dialog, 0, 0, 236, 110, "RPPH error detected"
  DropListBox 70, 50, 160, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", change_reason
  DropListBox 70, 65, 160, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", disenrollment_reason
  ButtonGroup ButtonPressed
    OkButton 125, 85, 50, 15
    CancelButton 180, 85, 50, 15
  Text 10, 55, 55, 10, "Change reason:"
  Text 10, 70, 60, 10, "Disenroll reason:"
  ButtonGroup ButtonPressed
    OkButton 155, 330, 50, 15
  Text 15, 20, 210, 10, "* Initial enrollment is selected, but has been enrolled previously"
  GroupBox 5, 5, 225, 40, "An error occurred on in RPPH. Typical errors include:"
  Text 15, 30, 210, 10, "* Exclusion code may be the same as the enrollment date"
EndDialog

BeginDialog excl_code_dialog, 0, 0, 191, 45, "Exclusion Code Error"
  ButtonGroup ButtonPressed
    OkButton 85, 25, 50, 15
    CancelButton 135, 25, 50, 15
  Text 15, 10, 155, 10, "Update the exclusion code field, then press OK."
EndDialog

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'call check_for_MMIS(True) 'Sending MMIS back to the beginning screen and checking for a password prompt
EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
IF MMIS_panel_check <> "RKEY" THEN 
	DO
		PF6
		EMReadScreen session_terminated_check, 18, 1, 7
	LOOP until session_terminated_check = "SESSION TERMINATED"
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit
	EMWriteScreen "x", 8, 3
	transmit
END IF 

'grabs the PMI number if one is listed on RKEY
EMReadscreen MMIS_case_number, 8, 9, 19
MMIS_case_number= trim(MMIS_case_number)

'do the dialog here
Do
	Do
		Dialog case_dlg
		cancel_confirmation
		If MMIS_case_number = "" then MsgBox "You must have a Case number to continue!"
		If health_plan = "Select one..." then MsgBox " You must select a health plan."
		If change_reason = "Select one..." then MsgBox " You must select a change reason."
		If Interpreter_yes = 1 and Interpreter_type = "Select one..." then MsgBox "You must select an interpreter language."
	Loop until Interpreter_yes = 0 or (Interpreter_yes = 1 and Interpreter_type <> "Select one...")
Loop until (MMIS_case_number <> "" and health_plan <> "Select one..." and change_reason <> "Select one...")

'blanking out varibles if the other option is selected
If change_reason = "Select one..." then change_reason = ""
If disenrollment_reason = "Select one..." then disenrollment_reason = ""

If Insurance_yes = checked then 
	insurance_yn = "Y"
Else 
	insurance_yn = "N"
End If 

If foster_care_yes = checked Then 
	foster_care_yn = "Y"
Else 
	foster_care_yn = "N"
End if 

'checking for an active MMIS session
Call check_for_MMIS(True)
EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
IF MMIS_panel_check <> "RKEY" THEN 
	DO
		PF6
		EMReadScreen session_terminated_check, 18, 1, 7
	LOOP until session_terminated_check = "SESSION TERMINATED"
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit
	EMWriteScreen "x", 8, 3
	transmit
END IF 

'formatting variables----------------------------------------------------------------------------------------------------
If len(enrollment_month) = 1 THEN enrollment_month = "0" & enrollment_month
IF len(enrollment_year) <> 2 THEN enrollment_year = right(enrollment_year, 2)

MNSURE_Case = False
If len(MMIS_case_number) = 8 AND left(MMIS_case_number, 1) <> 0 THEN MNSURE_Case = TRUE
MMIS_case_number = right("00000000" & MMIS_case_number, 8)

'MsgBox "MNSure Case? " & MNSURE_Case & vbNewLine & MMIS_case_number
enrollment_date = enrollment_month & "/01/" & enrollment_year

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "i", 2, 19
EMWriteScreen MMIS_case_number, 9, 19 
transmit
transmit
transmit
EMReadscreen RCIN_check, 4, 1, 49
If RCIN_check <> "RCIN" then script_end_procedure("The listed Case number was not found. Check your Case number and try again.")

Dim listed_clients_array
ReDim listed_clients_array (0)


rcin_row = 11
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen pmi_nbr, 8, rcin_row, 4
	EMReadscreen last_name, 17, rcin_row, 24
	EMReadscreen first_name, 9, rcin_row, 42
	last_name = trim(last_name)
	first_name = trim(first_name)
	client_string = pmi_nbr & " - " & last_name & ", " & first_name 
	client_array = client_array & client_string & "|"
	rcin_row = rcin_row + 1 
	If rcin_row = 21 Then 
		PF8
		EMReadScreen end_rcin, 6, 24, 2
		If end_rcin = "CANNOT" then Exit Do
		rcin_row = 11
	End If
	Emreadscreen last_clt_check, 8, rcin_row, 4
LOOP until last_clt_check = "        "			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

DIM all_client_array()
ReDim all_clients_array(total_clients, 1)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(x, 0) = Interim_array(x)
	PF7
	PF7
	Do
		Search_pmi = left(Interim_array(x), 8)
		row = 1 
		col = 1 
		EMSearch Search_pmi, row, col 
		If row = 0 then 
			PF8
		Else 
			EMReadScreen hc_status, 1, row, 76
			If hc_status = "A" Then all_clients_array(x, 1) = 1
			Exit Do
		End If 
		EMReadScreen end_rcin, 6, 24, 2
	Loop until end_rcin = "CANNOT"
NEXT



BEGINDIALOG HH_memb_dialog, 0, 0, 250, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 105, 10, "Household members to look at:"
	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 175, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
	ButtonGroup ButtonPressed
	OkButton 195, 10, 50, 15
	CancelButton 195, 30, 50, 15
ENDDIALOG

'runs the dialog that has been dynamically created. Streamlined with new functions.
Dialog HH_memb_dialog
If buttonpressed = 0 then stopscript

HH_member_array = ""

FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			'msgbox all_clients_
			HH_member_array = HH_member_array & left(all_clients_array(i, 0), 8) & " "
		END IF
	END IF
NEXT

HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
HH_member_array = SPLIT(HH_member_array, " ")

const client_name  = 0
const client_pmi   = 1
const current_plan = 2
const new_plan     = 3
const change_rsn   = 4
const disenrol_rsn = 5
const med_code     = 6 
const dent_code    = 7
const contr_code   = 8
const preg_yes 	   = 9
const interp_code  = 10 
const enrol_sucs   = 11

Dim MMIS_clients_array
ReDim MMIS_clients_array (12, 0)

EMReadScreen RCIN_check, 4, 1, 49
If RCIN_check = "RCIN" Then PF6 
EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
IF MMIS_panel_check <> "RKEY" THEN 
	DO
		PF6
		EMReadScreen session_terminated_check, 18, 1, 7
	LOOP until session_terminated_check = "SESSION TERMINATED"
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit
	EMWriteScreen "x", 8, 3
	transmit
END IF 

item = 0 

For each member in HH_member_array
	ReDim Preserve MMIS_clients_array(12, item)
	EMWriteScreen "I", 2, 19
	EMWriteScreen member, 4, 19 
	EMWriteScreen "        ", 9, 19
	transmit
	MMIS_clients_array (client_pmi, item) = member
	EMReadScreen last_name, 18, 3, 2
	EMReadScreen first_name, 12, 3, 20
	last_name = trim(last_name)
	first_name = trim(first_name)
	MMIS_clients_array (client_name, item) = last_name & ", " & first_name
	
	'check RPOL to see if there is other insurance available, if so worker processes manually
	'EMWriteScreen "X", 11, 2
	'Transmit 
	EMWriteScreen "RPOL", 1, 8
	transmit
	'making sure script got to right panel
	EMReadScreen RPOL_check, 4, 1, 52
	If RPOL_check <> "RPOL" then script_end_procedure("The script was unable to navigate to RPOL process manually if needed.")
	EMreadscreen policy_number, 1, 7, 8
	if policy_number <> " " then 
		PF6
		script_end_procedure ("This case has spans on RPOL. Please evaluate manually at this time.")
	end if
	PF6
	
	EMWriteScreen "RPPH", 1, 8
	transmit
	row = 1
	col = 1 
	EMSearch "99/99/99", row, col
	IF row < 10 Then 
		If col = 18 Then 
			EMReadScreen excl_code, 2, row, 2
		ElseIf col = 45 Then 
			EMReadScreen excl_code, 2, row, 29
		ElseIf col = 72 Then 
			EMReadScreen excl_code, 2, row, 56
		End If 
		If excl_code = "AA" Then MMIS_clients_array(current_plan, item) = "XCL - Adoption Assistance"
		If excl_code = "AB" Then MMIS_clients_array(current_plan, item) = "XCL - Part A or B Only"
		If excl_code = "BB" Then MMIS_clients_array(current_plan, item) = "XCL - Blind/Disabled under 65 years"
		If excl_code = "CC" Then MMIS_clients_array(current_plan, item) = "XCL - Child Protection Case"
		If excl_code = "CD" Then MMIS_clients_array(current_plan, item) = "XCL - Chemical Dependant Pilot"
		If excl_code = "CS" Then MMIS_clients_array(current_plan, item) = "XCL - Condumer Support Grant"
		If excl_code = "CV" Then MMIS_clients_array(current_plan, item) = "XCL - Center for Victims of Torture"
		If excl_code = "DD" Then MMIS_clients_array(current_plan, item) = "XCL - Communicable Disease"
		If excl_code = "DO" Then MMIS_clients_array(current_plan, item) = "XCL - Diability Opt Out"
		If excl_code = "EE" Then MMIS_clients_array(current_plan, item) = "XCL - SED/SPMI"
		If excl_code = "FF" Then MMIS_clients_array(current_plan, item) = "XCL - Child in Foster Care"
		If excl_code = "GG" THen MMIS_clients_array(current_plan, item) = "XCL - Geographic Exclusion"
		If excl_code = "HH" Then MMIS_clients_array(current_plan, item) = "XCL - Private HMO Coverage"
		If excl_code = "II" Then MMIS_clients_array(current_plan, item) = "XCL - Breast/Cervical Cancer"
		If excl_code = "IP" THen MMIS_clients_array(current_plan, item) = "XCL - Insurance Pending"
		If excl_code = "KK" Then MMIS_clients_array(current_plan, item) = "XCL - Elderly Waiver"
		If excl_code = "LL" Then MMIS_clients_array(current_plan, item) = "XCL - Personal Care Attendent"
		If excl_code = "MD" Then MMIS_clients_array(current_plan, item) = "XCL - MA Delay"
		If excl_code = "MM" Then MMIS_clients_array(current_plan, item) = "XCL - Native American on Reservation"
		If excl_code = "MS" Then MMIS_clients_array(current_plan, item) = "XCL - MNSURE Tracking"
		If excl_code = "PC" Then MMIS_clients_array(current_plan, item) = "XCL - Payment County"
		If excl_code = "QQ" Then MMIS_clients_array(current_plan, item) = "XCL - QMB/SLMB Eligibility"
		If excl_code = "RR" Then MMIS_clients_array(current_plan, item) = "XCL - Refugee/EMA/EGA"
		If excl_code = "SS" Then MMIS_clients_array(current_plan, item) = "XCL - Medical Spenddown"
		If excl_code = "TT" Then MMIS_clients_array(current_plan, item) = "XCL - Terminal Illness"
		If excl_code = "UU" Then MMIS_clients_array(current_plan, item) = "XCL - Limited Disability"
		If excl_code = "WW" Then MMIS_clients_array(current_plan, item) = "XCL - Delayed Nursing Home"
		If excl_code = "YY" Then MMIS_clients_array(current_plan, item) = "XCL - Delayed Decision"
		If excl_code = "ZZ" Then MMIS_clients_array(current_plan, item) = "XCL - RTC/IMD Resident"
	Else 
		EMReadScreen hp_code, 10, row, 23
		
		If hp_code = "A585713900" then MMIS_clients_array(current_plan, item) = "Health Partners" 
		If hp_code = "A565813600" then MMIS_clients_array(current_plan, item) = "Ucare"
		If hp_code = "A405713900" then MMIS_clients_array(current_plan, item) = "Medica" 
		If hp_code = "A065813800" then MMIS_clients_array(current_plan, item) = "Blue Plus"
		If hp_code = "A836618200" then MMIS_clients_array(current_plan, item) = "Hennepin Health PMAP"
		If hp_code = "A965713400" then MMIS_clients_array(current_plan, item) = "Hennepin Health SNBC"
	End If
	MMIS_clients_array(new_plan,     item) = health_plan
	MMIS_clients_array(change_rsn,   item) = change_reason
	MMIS_clients_array(disenrol_rsn, item) = disenrollment_reason
	PF6
	EMWaitReady 0, 0
	item = item + 1 
Next

x = 0
max = Ubound(MMIS_clients_array, 2)

BeginDialog Enrollment_dlg, 0, 0, 750, (max * 20) + 60, "Enrollment Information"
  Text 5, 5, 25, 10, "Name"
  Text 100, 5, 15, 10, "PMI"
  Text 145, 5, 75, 10, "Current Plan/Exclusion"
  Text 250, 5, 50, 10, "Medical Clinic"
  Text 310, 5, 45, 10, "Dental Clinic"
  Text 370, 5, 40, 10, "Health plan:"
  Text 440, 5, 55, 10, "Contract Code:"
  Text 500, 5, 55, 10, "Change reason:"
  Text 565, 5, 60, 10, "Disenroll reason:"
  Text 640, 5, 35, 10, "Pregnant?"
  Text 695, 5, 55, 10, "Interpreter Code"

  For person = 0 to Ubound(MMIS_clients_array, 2)
  	Text 5, (x * 20) + 25, 95, 10, MMIS_clients_array(client_name, person)
  	Text 100, (x * 20) + 25, 35, 10, MMIS_clients_array(client_pmi, person)
  	Text 145, (x * 20) + 25, 95, 10, MMIS_clients_array(current_plan, person)
  	EditBox 250, (x * 20) + 20, 55, 15, MMIS_clients_array(med_code, person)  
  	EditBox 310, (x * 20) + 20, 50, 15, MMIS_clients_array(dent_code, person)
    DropListBox 370, (x * 20) + 20, 60, 15, " "+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Hennepin Health SNBC"+chr(9)+"Ucare", MMIS_clients_array(new_plan, person)
  	DropListBox 440, (x * 20) + 20, 50, 15, "MA 12"+chr(9)+"NM 12"+chr(9)+"MA 30"+chr(9)+"MA 35", MMIS_clients_array(contr_code, person)
	DropListBox 500, (x * 20) + 20, 60, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", MMIS_clients_array(change_rsn, person)
  	DropListBox 565, (x * 20) + 20, 60, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", MMIS_clients_array(disenrol_rsn, person)
  	CheckBox 645, (x * 20) + 20, 25, 10, "Yes", MMIS_clients_array(preg_yes, person)
	EditBox 700, (x * 20) + 20, 25, 15, MMIS_clients_array(interp_code, person)
	x = x + 1
  Next

  Text 445, (max * 20) + 45, 60, 10, "Worker Signature"
  EditBox 510, (max * 20) + 40, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 640, (max * 20) + 40, 50, 15
    CancelButton 695, (max * 20) + 40, 50, 15
EndDialog

Do 
	Dialog Enrollment_dlg
	cancel_confirmation
Loop Until ButtonPressed = OK

process_manually_message = ""

If MNSURE_Case = TRUE Then 
	For member = 0 to Ubound(MMIS_clients_array, 2)
		'MMIS Codes
		Enrollment_date = Enrollment_month & "/01/" & enrollment_year
		'change reasons
		If MMIS_clients_array(change_rsn, member) = "First year change option" 	then change_reason = "FY"
		If MMIS_clients_array(change_rsn, member) = "Health plan contract end" 	then change_reason = "HP"
		If MMIS_clients_array(change_rsn, member) = "Initial enrollment"       	then change_reason = "IN"
		If MMIS_clients_array(change_rsn, member) = "Move"                     	then change_reason = "MV"
		If MMIS_clients_array(change_rsn, member) = "Ninety Day change option" 	then change_reason = "NT"
		If MMIS_clients_array(change_rsn, member) = "Open enrollment"    	  	then change_reason = "OE"
		If MMIS_clients_array(change_rsn, member) = "PMI merge" 				then change_reason = "PM"
		If MMIS_clients_array(change_rsn, member) = "Reenrollment" 			  	then change_reason = "RE"
		If MMIS_clients_array(change_rsn, member) = "Select one..." 			then change_reason = ""
		
		'Disenrollment reasons
		If MMIS_clients_array(disenrol_rsn, member) = "Eligibility ended"        then disenrollment_reason = "EE"
		If MMIS_clients_array(disenrol_rsn, member) = "Exclusion"                then disenrollment_reason = "EX"
		If MMIS_clients_array(disenrol_rsn, member) = "First year change option" then disenrollment_reason = "FY"
		If MMIS_clients_array(disenrol_rsn, member) = "Health plan contract end" then disenrollment_reason = "HP"
		If MMIS_clients_array(disenrol_rsn, member) = "Jail - Incarceration"     then disenrollment_reason = "JL"
		If MMIS_clients_array(disenrol_rsn, member) = "Move"                     then disenrollment_reason = "MV"
		If MMIS_clients_array(disenrol_rsn, member) = "Loss of disability"       then disenrollment_reason = "ND"
		If MMIS_clients_array(disenrol_rsn, member) = "Ninety Day change option" then disenrollment_reason = "NT"
		If MMIS_clients_array(disenrol_rsn, member) = "Open Enrollment"          then disenrollment_reason = "OE"
		If MMIS_clients_array(disenrol_rsn, member) = "PMI merge"                then disenrollment_reason = "PM"
		If MMIS_clients_array(disenrol_rsn, member) = "Voluntary"                then disenrollment_reason = "VL"
		If MMIS_clients_array(disenrol_rsn, member) = "Select one..."            then disenrollment_reason = ""
		
		'REFM Codes
		If MMIS_clients_array(preg_yes, member) = checked Then 
			pregnant_yn = "Y"
		Else
			pregnant_yn = "N"
		End If 
		
		If MMIS_clients_array(interp_code, member) = "" Then 
			interpreter_yn = "N"
		Else 
			interpreter_yn = "Y"
		End If 
		
		EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
		IF MMIS_panel_check <> "RKEY" THEN 
			DO
				PF6
				EMReadScreen session_terminated_check, 18, 1, 7
			LOOP until session_terminated_check = "SESSION TERMINATED"
			'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
			EMWriteScreen "mw00", 1, 2
			transmit
			transmit
			EMWriteScreen "x", 8, 3
			transmit
		END IF 
		'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
		EMWriteScreen "c", 2, 19
		EMWriteScreen "        ", 9, 19
		EMWriteScreen MMIS_clients_array(client_pmi, member), 4, 19 
		transmit
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine

		DO
			'check RPOL to see if there is other insurance available, if so worker processes manually
			EMWriteScreen "rpol", 1, 8
			transmit
			'making sure script got to right panel
			EMReadScreen RPOL_check, 4, 1, 52
			If RPOL_check <> "RPOL" then process_manually_message = process_manually_message & "Could not navigate to RPOL for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
			EMreadscreen policy_number, 1, 7, 8
			If policy_number <> " " then process_manually_message = process_manually_message & "RPOL for PMI " & MMIS_clients_array(client_pmi, member) & " has a span listed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine

			'nav to RPPH
			EMWriteScreen "rpph", 1, 8
			transmit

			'making sure script got to right panel
			EMReadScreen RPPH_check, 4, 1, 52
			If RPPH_check <> "RPPH" then process_manually_message = process_manually_message & "Could not navigate to RPPH for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
	 
			Enrollment_date = Enrollment_month & "/01/" & enrollment_year
			xcl_end_date = DateAdd("d", -1, enrollment_date)
			xcl_end_month = right("00" & DatePart("m", xcl_end_date), 2)
			xcl_end_day   = right("00" & DatePart("d", xcl_end_date), 2)
			xcl_end_year  = right(DatePart("yyyy", xcl_end_date), 2)
			xcl_end_date  = xcl_end_month & "/" & xcl_end_day & "/" & xcl_end_year
''			msgbox enrollment_date & vbNewLine & xcl_end_date
			'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
			If left(MMIS_clients_array(current_plan, member), 3) = "XCL" Then 
				If MMIS_clients_array(current_plan, member) = "XCL - Delayed Decision" Then 
					row = 1
					col = 1 
					EMSearch "99/99/99", row, col
''					msgbox "Row: " & row & vbNewLine & "Col: " & col
					If col <> 0 Then EMWriteScreen xcl_end_date, row, col
				Else 
					process_manually_message = process_manually_message & "There is an exclusion code other than 'YY' for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				End If 
			End If 
			EMReadscreen XCL_code, 2, 6, 2
			If XCL_code = "* " then
				EMSetCursor 6, 2
				EMSendKey "..."
			End if
''			msgbox "Exclusion ended"

			If MMIS_clients_array(new_plan, member) = "Health Partners" then health_plan_code = "A585713900"
			If MMIS_clients_array(new_plan, member) = "Ucare" then health_plan_code = "A565813600"
			If MMIS_clients_array(new_plan, member) = "Medica" then health_plan_code = "A405713900"
			If MMIS_clients_array(new_plan, member) = "Blue Plus" then health_plan_code = "A065813800"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health PMAP" then health_plan_code = "A836618200"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health SNBC" then health_plan_code = "A965713400"
			
			contract_code = MMIS_clients_array (contr_code, member)
			Contract_code_part_one = left(contract_code, 2)
			Contract_code_part_two = right(contract_code, 2)
			
			'enter disenrollment reason
			If change_reason <> "" Then 
				EMWriteScreen disenrollment_reason, 14, 75
			Else
				EMWriteScreen disenrollment_reason, 13, 75
			End If 
			
			'resets to bottom of the span list. 
			pf11

			'enter enrollment date
			EMWriteScreen enrollment_date, 13, 5
			'enter managed care plan code
			EMWriteScreen health_plan_code, 13, 23
			'enter contract code
			EMWriteScreen contract_code_part_one, 13, 34
			EMWriteScreen contract_code_part_two, 13, 37
			'enter change reason
			EMWriteScreen change_reason, 13, 71

			EMWaitReady 0, 0
			
			EMReadScreen false_end, 8, 14, 14
			If false_end = "99/99/99" Then 
				EMReadScreen double_check, 2, 14, 5
				If double_check = "  " Then EMWriteScreen "...", 14, 5
			End If 
			
''			msgbox "RPPH updated"

			'REFM screen
			EMWriteScreen "refm", 1, 8
			transmit
			EMReadScreen RPPH_error_check, 10, 24, 2
			If trim(RPPH_error_check) = "EXCLSN END" then 
				Do
					Dialog excl_code_dialog
					cancel_confirmation
					transmit
					EMReadScreen RPPH_error_check, 10, 24, 2
				Loop until trim(RPPH_error_check) <> "EXCLSN END" 
''				Msgbox "Updated the exclusion code field, then press OK."
''				transmit
			ELSEIF trim(RPPH_error_check) <> "" then 
				dialog RPPH_error_dialog
				If buttonpressed = 0 then script_end_procedure("Error message was not resolved. Please review enrollment information before trying the script again.")
				EMWriteScreen "...", 13, 5
				EMReadScreen false_end, 8, 14, 14
				If false_end = "99/99/99" Then 
					EMReadScreen double_check, 2, 14, 5
					If double_check = "??" Then EMWriteScreen "...", 14, 5
				End If 
			END IF 
	        
	        'blanking out varibles if the other option is selected
	        If change_reason = "Select one..." then change_reason = ""
	        If disenrollment_reason = "Select one..." then disenrollment_reason = ""
			'making sure script got to right panel
			EMReadScreen REFM_check, 4, 1, 52
			If REFM_check <> "REFM" then process_manually_message = process_manually_message & "The script was unable to navigate to REFM for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
		Loop until REFM_check = "REFM"
		
		'form rec'd
		EMsetcursor 10, 16
		EMSendkey "y"
		'other insurance y/n
		EMsetcursor 11, 18
		EMsendkey insurance_yn
		'preg y/n
		EMsetcursor 12, 19
		EMsendkey pregnant_yn
		'interpreter y/n
		EMsetcursor 13, 29
		EMsendkey interpreter_yn
		'interpreter type
		if MMIS_clients_array(interp_code, member) <> "" then
			EMsetcursor 13, 52
			EMsendKey MMIS_clients_array(interp_code, member)
		end if
		'medical clinic code
		EMsetcursor 19, 4
		EMsendkey MMIS_clients_array(med_code, member)
		'dental clinic code if applicable
		EMsetcursor 19, 24
		EMsendkey MMIS_clients_array(dent_code, member)
		'foster care y/n
		EMsetcursor 21, 15
		EMsendkey foster_care_yn
''		msgbox "REFM updated"
		PF9

		'error handling to ensure that enrollment date and exclusion dates don't conflict
		EMReadScreen REFM_error_check, 19, 24, 2 'checks for an inhibiting edit
		If enrollment_year < "16" AND REFM_error_check = "WARNING: MA12,01/16" Then
			script_end_procedure("This health plan is not available until 01/01/16." & vbNewLine & "Make sure you change the enrollment date when using the script again.")
		ELSEIF REFM_error_check <> "WARNING: MA12,01/16" Then
			IF REFM_error_check <> "                   " then
                IF REFM_error_check <> "INVALID KEY ENTERED" then 
                    EMReadScreen full_error_msg, 79, 24, 2
                    full_error_msg = trim(full_error_msg)
				    process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & full_error_msg & vbNewLine & vbNewLine
			    END if 
            END IF 
		END IF 
		'Save and case note
		pf3
		EMWriteScreen "c", 2, 19
		transmit
		EMReadScreen rsum_enrollment, 8, 16, 20
		EMReadScreen rsum_plan, 10, 16, 52
''		MsgBox "RSUM date and plan: " & rsum_enrollment & " - " & rsum_plan & vbNewLine & "Coded date and plan: " & Enrollment_date & " - " & health_plan_code
		IF rsum_enrollment = Enrollment_date AND rsum_plan = health_plan_code Then
			MMIS_clients_array(enrol_sucs, member) = TRUE
''			pf4
''			pf11
''			EMSendkey "***HMO Note*** " &  MMIS_clients_array(client_name, member) & " enrolled into " &  MMIS_clients_array(new_plan, member) & " " & Enrollment_date & " " & worker_signature
''			pf3
		Else 
			failed_enrollment_message = failed_enrollment_message & vbNewLine & vbNewLine & process_manually_message
			MMIS_clients_array(enrol_sucs, member) = FALSE		
		End If 
''		MsgBox process_manually_message
		pf3
		IF REFM_error_check = "WARNING: MA12,01/16" Then
			PF3
		END IF
		process_manually_message = ""
	Next 
Else
	For member = 0 to Ubound(MMIS_clients_array, 2)
		'MMIS Codes
		'change reasons
		If MMIS_clients_array(change_rsn, member) = "First year change option" 	then change_reason = "FY"
		If MMIS_clients_array(change_rsn, member) = "Health plan contract end" 	then change_reason = "HP"
		If MMIS_clients_array(change_rsn, member) = "Initial enrollment"       	then change_reason = "IN"
		If MMIS_clients_array(change_rsn, member) = "Move"                     	then change_reason = "MV"
		If MMIS_clients_array(change_rsn, member) = "Ninety Day change option" 	then change_reason = "NT"
		If MMIS_clients_array(change_rsn, member) = "Open enrollment"    	  	then change_reason = "OE"
		If MMIS_clients_array(change_rsn, member) = "PMI merge" 				then change_reason = "PM"
		If MMIS_clients_array(change_rsn, member) = "Reenrollment" 			  	then change_reason = "RE"
		If MMIS_clients_array(change_rsn, member) = "Select one..." 			then change_reason = ""
		
		'Disenrollment reasons
		If MMIS_clients_array(disenrol_rsn, member) = "Eligibility ended"        then disenrollment_reason = "EE"
		If MMIS_clients_array(disenrol_rsn, member) = "Exclusion"                then disenrollment_reason = "EX"
		If MMIS_clients_array(disenrol_rsn, member) = "First year change option" then disenrollment_reason = "FY"
		If MMIS_clients_array(disenrol_rsn, member) = "Health plan contract end" then disenrollment_reason = "HP"
		If MMIS_clients_array(disenrol_rsn, member) = "Jail - Incarceration"     then disenrollment_reason = "JL"
		If MMIS_clients_array(disenrol_rsn, member) = "Move"                     then disenrollment_reason = "MV"
		If MMIS_clients_array(disenrol_rsn, member) = "Loss of disability"       then disenrollment_reason = "ND"
		If MMIS_clients_array(disenrol_rsn, member) = "Ninety Day change option" then disenrollment_reason = "NT"
		If MMIS_clients_array(disenrol_rsn, member) = "Open Enrollment"          then disenrollment_reason = "OE"
		If MMIS_clients_array(disenrol_rsn, member) = "PMI merge"                then disenrollment_reason = "PM"
		If MMIS_clients_array(disenrol_rsn, member) = "Voluntary"                then disenrollment_reason = "VL"
		If MMIS_clients_array(disenrol_rsn, member) = "Select one..."            then disenrollment_reason = ""
		
		'REFM Codes
		If MMIS_clients_array(preg_yes, member) = checked Then 
			pregnant_yn = "Y"
		Else
			pregnant_yn = "N"
		End If 
		
		If MMIS_clients_array(interp_code, member) = "" Then 
			interpreter_yn = "N"
		Else 
			interpreter_yn = "Y"
		End If 	
		
		EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
		IF MMIS_panel_check <> "RKEY" THEN 
			DO
				PF6
				EMReadScreen session_terminated_check, 18, 1, 7
			LOOP until session_terminated_check = "SESSION TERMINATED"
			'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
			EMWriteScreen "mw00", 1, 2
			transmit
			transmit
			EMWriteScreen "x", 8, 3
			transmit
		END IF 
''		msgbox "At RKEY"
		'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
		EMWriteScreen "c", 2, 19
		EMWriteScreen "        ", 4, 19
		EMWriteScreen MMIS_case_number, 9, 19 
		transmit
		transmit
		transmit
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
		Do
			row = 1
			col = 1 
			EMSearch MMIS_clients_array(client_pmi, member), row, col
			If row = 0 Then
				PF8 
				EMReadScreen end_of_clts, 6, 24, 2
				If end_of_clts = "CANNOT" Then 
					process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be found on this case. The enrollment for " &  MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
					Exit Do
				End If 
			End If	
		Loop until row <> 0
		EMWriteScreen "X", row, 2
''		msgbox "person selected"
		transmit
''		msgbox "at RSUM"
		EMReadscreen RKEY_check, 4, 1, 52
		If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
''		msgbox process_manually_message
		
		DO
			'check RPOL to see if there is other insurance available, if so worker processes manually
			EMWriteScreen "rpol", 1, 8
			transmit
			'making sure script got to right panel
			EMReadScreen RPOL_check, 4, 1, 52
			If RPOL_check <> "RPOL" then process_manually_message = process_manually_message & "Could not navigate to RPOL for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
			EMreadscreen policy_number, 1, 7, 8
			If policy_number <> " " then process_manually_message = process_manually_message & "RPOL for PMI " & MMIS_clients_array(client_pmi, member) & " has a span listed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine

			'nav to RPPH
			EMWriteScreen "rpph", 1, 8
			transmit

			'making sure script got to right panel
			EMReadScreen RPPH_check, 4, 1, 52
			If RPPH_check <> "RPPH" then process_manually_message = process_manually_message & "Could not navigate to RPPH for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine

			Enrollment_date = Enrollment_month & "/01/" & enrollment_year
			xcl_end_date = DateAdd("d", -1, enrollment_date)
			xcl_end_month = right("00" & DatePart("m", xcl_end_date), 2)
			xcl_end_day   = right("00" & DatePart("d", xcl_end_date), 2)
			xcl_end_year  = right(DatePart("yyyy", xcl_end_date), 2)
			xcl_end_date  = xcl_end_month & "/" & xcl_end_day & "/" & xcl_end_year
''			msgbox enrollment_date & vbNewLine & xcl_end_date
			'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
			If left(MMIS_clients_array(current_plan, member), 3) = "XCL" Then 
				If MMIS_clients_array(current_plan, member) = "XCL - Delayed Decision" Then 
					row = 1
					col = 1 
					EMSearch "99/99/99", row, col
''					msgbox "Row: " & row & vbNewLine & "Col: " & col
					If col <> 0 Then EMWriteScreen xcl_end_date, row, col
				Else 
					process_manually_message = process_manually_message & "There is an exclusion code other than 'YY' for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				End If 
			End If 
			EMReadscreen XCL_code, 2, 6, 2
			If XCL_code = "* " then
				EMSetCursor 6, 2
				EMSendKey "..."
			End if
''			msgbox "Exclusion ended"

			If MMIS_clients_array(new_plan, member) = "Health Partners" then health_plan_code = "A585713900"
			If MMIS_clients_array(new_plan, member) = "Ucare" then health_plan_code = "A565813600"
			If MMIS_clients_array(new_plan, member) = "Medica" then health_plan_code = "A405713900"
			If MMIS_clients_array(new_plan, member) = "Blue Plus" then health_plan_code = "A065813800"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health PMAP" then health_plan_code = "A836618200"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health SNBC" then health_plan_code = "A965713400"
			
			contract_code = MMIS_clients_array (contr_code, member)
			Contract_code_part_one = left(contract_code, 2)
			Contract_code_part_two = right(contract_code, 2)
			
			'enter disenrollment reason
			If change_reason <> "" Then 
				EMWriteScreen disenrollment_reason, 14, 75
			Else
				EMWriteScreen disenrollment_reason, 13, 75
			End If 
			
			'resets to bottom of the span list. 
			pf11

			'enter enrollment date
			EMWriteScreen enrollment_date, 13, 5
			'enter managed care plan code
			EMWriteScreen health_plan_code, 13, 23
			'enter contract code
			EMWriteScreen contract_code_part_one, 13, 34
			EMWriteScreen contract_code_part_two, 13, 37
			'enter change reason
			EMWriteScreen change_reason, 13, 71

			EMWaitReady 0, 0
			
			EMReadScreen false_end, 8, 14, 14
			If false_end = "99/99/99" Then 
				EMReadScreen double_check, 2, 14, 5
				If double_check = "  " Then EMWriteScreen "...", 14, 5
			End If 
			'msgbox "RPPH updated"

			'REFM screen
			EMWriteScreen "refm", 1, 8
			transmit
			EMReadScreen RPPH_error_check, 10, 24, 2
			If trim(RPPH_error_check) = "EXCLSN END" then 
				Do
					Dialog excl_code_dialog
					cancel_confirmation
					transmit
					EMReadScreen RPPH_error_check, 10, 24, 2
				Loop until trim(RPPH_error_check) <> "EXCLSN END" 
''				Msgbox "Updated the exclusion code field, then press OK."
''				transmit
			ELSEIF trim(RPPH_error_check) <> "" then 
				dialog RPPH_error_dialog
				If buttonpressed = 0 then script_end_procedure("Error message was not resolved. Please review enrollment information before trying the script again.")
				EMWriteScreen "...", 13, 5
			END IF 
			
			'blanking out varibles if the other option is selected
			If change_reason = "Select one..." then change_reason = ""
			If disenrollment_reason = "Select one..." then disenrollment_reason = ""
			'making sure script got to right panel
			EMReadScreen REFM_check, 4, 1, 52
			If REFM_check <> "REFM" then process_manually_message = process_manually_message & "The script was unable to navigate to REFM for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
		Loop until REFM_check = "REFM"

		'form rec'd
		EMsetcursor 10, 16
		EMSendkey "y"
		'other insurance y/n
		EMsetcursor 11, 18
		EMsendkey insurance_yn
		'preg y/n
		EMsetcursor 12, 19
		EMsendkey pregnant_yn
		'interpreter y/n
		EMsetcursor 13, 29
		EMsendkey interpreter_yn
		'interpreter type
		if MMIS_clients_array(interp_code, member) <> "" then
			EMsetcursor 13, 52
			EMsendKey MMIS_clients_array(interp_code, member)
		end if
		'medical clinic code
		EMsetcursor 19, 4
		EMsendkey MMIS_clients_array(med_code, member)
		'dental clinic code if applicable
		EMsetcursor 19, 24
		EMsendkey MMIS_clients_array(dent_code, member)
		'foster care y/n
		EMsetcursor 21, 15
		EMsendkey foster_care_yn
''		msgbox "REFM updated"
		PF9

		'error handling to ensure that enrollment date and exclusion dates don't conflict
		EMReadScreen REFM_error_check, 19, 24, 2 'checks for an inhibiting edit
		If enrollment_year < "16" AND REFM_error_check = "WARNING: MA12,01/16" Then
			script_end_procedure("This health plan is not available until 01/01/16." & vbNewLine & "Make sure you change the enrollment date when using the script again.")
		ELSEIF REFM_error_check <> "WARNING: MA12,01/16" Then
			IF REFM_error_check <> "                   " then
                IF REFM_error_check <> "INVALID KEY ENTERED" then 
                    EMReadScreen full_error_msg, 79, 24, 2
                    full_error_msg = trim(full_error_msg)
				    process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & full_error_msg & vbNewLine & vbNewLine
			    END IF 
            END IF 
		END IF 
''		msgbox "all updated - see casenote code"
		'Save and case note
		pf3
		EMWriteScreen "c", 2, 19
		transmit
		EMReadScreen rsum_enrollment, 8, 16, 20
		EMReadScreen rsum_plan, 10, 16, 52
''		MsgBox "RSUM date and plan: " & rsum_enrollment & " - " & rsum_plan & vbNewLine & "Coded date and plan: " & Enrollment_date & " - " & health_plan_code
		IF rsum_enrollment = Enrollment_date AND rsum_plan = health_plan_code Then
			MMIS_clients_array(enrol_sucs, member) = TRUE
		''			pf4
		''			pf11
		''			EMSendkey "***HMO Note*** " &  MMIS_clients_array(client_name, member) & " enrolled into " &  MMIS_clients_array(new_plan, member) & " " & Enrollment_date & " " & worker_signature
		''			pf3
		Else 
			failed_enrollment_message = failed_enrollment_message & vbNewLine & vbNewLine & process_manually_message
			MMIS_clients_array(enrol_sucs, member) = FALSE		
		End If 
''		MsgBox process_manually_message
		pf3
		IF REFM_error_check = "WARNING: MA12,01/16" Then
			PF3
		END IF
		process_manually_message = ""
	Next
End If 

EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
IF MMIS_panel_check <> "RKEY" THEN 
	DO
		PF6
		EMReadScreen session_terminated_check, 18, 1, 7
	LOOP until session_terminated_check = "SESSION TERMINATED"
	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit
	EMWriteScreen "x", 8, 3
	transmit
END IF 

'Case Noting - goes into RSUM for the first client to do the case note
EMWriteScreen "c", 2, 19
EMWriteScreen "        ", 9, 19
EMWriteScreen MMIS_clients_array(client_pmi, 0), 4, 19 
transmit
pf4
pf11		'Starts a new case note'

EMWriteScreen "***HMO Note*** Household enrollment updated for " & Enrollment_date, 5, 8
row = 6
For member = 0 to Ubound(MMIS_clients_array, 2)
	If MMIS_clients_array(enrol_sucs, member) = TRUE Then
		EMWriteScreen MMIS_clients_array(client_name, member) & " enrolled into " & MMIS_clients_array(new_plan, member), row, 8
		row = row + 1
	End If 
Next 
EMWriteScreen "Processed by " & worker_signature, row, 8
pf3
pf3
IF REFM_error_check = "WARNING: MA12,01/16" Then
	PF3
END IF


failed_enrollment_message = "The script is complete. Enrollment has been updated and case noted." & vbNewLine & "There may be some clients enrollments that could not be processed by the script for some reason, they will be listed below:" & vbNewLine & "*****" & vbNewLine & vbNewLine & failed_enrollment_message

script_end_procedure (failed_enrollment_message)