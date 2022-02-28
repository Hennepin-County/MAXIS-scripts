'**********THIS IS A HENNEPIN SPECIFIC SCRIPT.  IF YOU REVERSE ENGINEER THIS SCRIPT, JUST BE CAREFUL.************
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MHC ENROLLMENT.vbs"
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
call changelog_update("02/18/2022", "Medica Plan will no longer default to the contract code of 'MA 30'##~## ##~##All plans will default to the contract code of 'MA 12' and if this selection is not changed manually during the script run, the script will enter the enrollment(s) as 'MA 12'.##~##", "Casey Love, Hennepin County")
call changelog_update("11/04/2021", "Open Enrollment Dates updated for 2022.", "Casey Love, Hennepin County")
call changelog_update("11/03/2021", "Added United HealthCare plan option as a selection with the plan code. This plan is available starting in January 2022.##~## ##~##DO NOT ENROLL RESIDENTS IN THIS PLAN PRIOR TO 01/2022. MMIS WILL ERROR AS THAT PLAN IS NOT AVAILBLE FOR A MONTH PRIOR TO 01/2022##~##", "Casey Love, Hennepin County")
call changelog_update("04/28/2021", "New functionality added!##~## ##~##Script will now read RCAS for the SVC LOC to ensure the case is handled by Hennepin County. This functionality is BRAND NEW and in testing. Tell us is there are any issues (the script preventing you from working on any cases that you can update or not correctly finding cases that you CAN'T update.) We won't know if we are looking for the right information or evaluating it correctly until we try it out.##~##", "Casey Love, Hennepin County")
call changelog_update("04/26/2021", "Script can now end 'HH' Exclusions for PMAP.##~##", "Casey Love, Hennepin County")
call changelog_update("04/26/2021", "Phone Number field is changed to a COMBOBOX which will allow you to select a phone number from a list based on what has been entered in and listed on RCAD in MMIS. This field can still be typed in, to allow for entry of a number not known to the system.##~##", "Casey Love, Hennepin County")
call changelog_update("04/06/2021", "New 'disenrollment reason' option created to 'DELETE SPAN' which will allow the removal of an enrollment span using '...' to remove a span that starts in the same month as an enrollment you are trying to create.##~##", "Casey Love, Hennepin County")
call changelog_update("04/06/2021", "Added handling for discovering a failed enrollment and allowing for a change in selections.##~##", "Casey Love, Hennepin County")
call changelog_update("11/20/2020", "BUG FIXES AND UPDATES##~## ##~##1. Changed the NOTE so that it will not create a note if no one has actually been enrolled.##~##2. Adjusted the end script wording to be more specific about what happened. ##~##3. Changed the 'Is this Open Enrollment' question to only appear from October until the November Cutoff date. You should not see the question now until next October.##~##", "Casey Love, Hennepin County")
call changelog_update("10/06/2020", "Added phone number field to the dialog for when the enrollment is requested by phone.", "Casey Love, Hennepin County")
call changelog_update("09/11/2020", "This script now contains the functionality for Open Enrollment and for any other enrollment.##~## ##~##The seperate script for Open Enrollment will no longer be available.##~##", "Casey Love, Hennepin County")
call changelog_update("12/19/2019", "Added IM 12 as an option for contract codes.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Added the ability for the script to delete the current enrollment plan if the beginning date for the current plan is the same as the new enrollment date.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Added the ability for the script to delete the Delayed Decision exclusion if the start date is the same as the enrollment date.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Changed the script coding on REFM screen to enter 'N' if enrollment information source was NOT the Paper Enrollment Form.", "Casey Love, Hennepin County")
call changelog_update("04/17/2019", "Resolving a BUG for METS cases enrolling for the first time, no exclusion code is defaulted.", "Casey Love, Hennepin County")
call changelog_update("04/16/2019", "BUG when disenrolling and reenrolling in a different plan. Functionality should work to disenroll and renroll in the same run - specific to issues discovered with NT option.", "Casey Love, Hennepin County")
call changelog_update("04/02/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
script_run_lowdown = ""
Function read_detail_on_RPPH(curr_enrl_date, curr_hp_code, curr_cntrct_code_one, curr_cntrct_code_two, curr_cntrct_code_comb, curr_change_rsn_code, curr_disenrl_rsn_code, curr_hp_name, curr_chg_rsn_full, curr_disenrl_rsn_full)
	EMReadScreen curr_enrl_date, 8, 13, 5
	'enter managed care plan code
	EMReadScreen curr_hp_code, 10, 13, 23
	'enter contract code
	EMReadScreen curr_cntrct_code_one, 2, 13, 34
	EMReadScreen curr_cntrct_code_two, 2, 13, 37
	'enter change reason
	EMReadScreen curr_change_rsn_code, 2, 13, 71
	EMReadScreen curr_disenrl_rsn_code, 2, 14, 75

	MMIS_clients_array(manual_enrollment, member) = TRUE

	curr_cntrct_code_comb = curr_cntrct_code_one & " " & curr_cntrct_code_two

	If curr_hp_code = "A585713900" then curr_hp_name = "Health Partners"
	If curr_hp_code = "A565813600" then curr_hp_name = "Ucare"
	If curr_hp_code = "A405713900" then curr_hp_name = "Medica"
	If curr_hp_code = "A065813800" then curr_hp_name = "Blue Plus"
	If curr_hp_code = "A168407400" then curr_hp_name = "United Healthcare"
	If curr_hp_code = "A836618200" then curr_hp_name = "Hennepin Health PMAP"
	If curr_hp_code = "A965713400" then curr_hp_name = "Hennepin Health SNBC"
	If curr_change_rsn_code = "FY" Then curr_chg_rsn_full = "First year change option"
	If curr_change_rsn_code = "HP" Then curr_chg_rsn_full = "Health plan contract end"
	If curr_change_rsn_code = "IN" Then curr_chg_rsn_full = "Initial enrollment"
	If curr_change_rsn_code = "MV" Then curr_chg_rsn_full = "Move"
	If curr_change_rsn_code = "NT" Then curr_chg_rsn_full = "Ninety Day change option"
	If curr_change_rsn_code = "OE" Then curr_chg_rsn_full = "Open enrollment"
	If curr_change_rsn_code = "PM" Then curr_chg_rsn_full = "PMI merge"
	If curr_change_rsn_code = "RE" Then curr_chg_rsn_full = "Reenrollment"
	If curr_disenrl_rsn_code = "EE" Then curr_disenrl_rsn_full = "Eligibility ended"
	If curr_disenrl_rsn_code = "EX" Then curr_disenrl_rsn_full = "Exclusion"
	If curr_disenrl_rsn_code = "FY" Then curr_disenrl_rsn_full = "First year change option"
	If curr_disenrl_rsn_code = "HP" Then curr_disenrl_rsn_full = "Health plan contract end"
	If curr_disenrl_rsn_code = "JL" Then curr_disenrl_rsn_full = "Jail - Incarceration"
	If curr_disenrl_rsn_code = "MV" Then curr_disenrl_rsn_full = "Move"
	If curr_disenrl_rsn_code = "ND" Then curr_disenrl_rsn_full = "Loss of disability"
	If curr_disenrl_rsn_code = "NT" Then curr_disenrl_rsn_full = "Ninety Day change option"
	If curr_disenrl_rsn_code = "OE" Then curr_disenrl_rsn_full = "Open Enrollment"
	If curr_disenrl_rsn_code = "PM" Then curr_disenrl_rsn_full = "PMI merge"
	If curr_disenrl_rsn_code = "VL" Then curr_disenrl_rsn_full = "Voluntary"
End Function

function enter_detail_on_refm()
	If enrollment_source = "Paper Enrollment Form" Then
		'form rec'd
		EMWriteScreen "Y", 10, 16
		'other insurance y/n
		EMWriteScreen insurance_yn, 11, 18
		'preg y/n
		EMWriteScreen pregnant_yn, 12, 19
		'interpreter y/n
		EMWriteScreen interpreter_yn, 13, 29
		'interpreter type
		if MMIS_clients_array(interp_code, member) <> "" then
			EMWriteScreen MMIS_clients_array(interp_code, member), 13, 52
		end if
		'medical clinic code
		EMWriteScreen MMIS_clients_array(med_code, member), 19, 4
		'dental clinic code if applicable
		EMWriteScreen MMIS_clients_array(dent_code, member), 19, 24
		'foster care y/n
		EMWriteScreen foster_care_yn, 21, 15
		' msgbox "REFM updated"
	Else
		'form rec'd
		EMWriteScreen "N", 10, 16
	End If
End Function

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

'call check_for_MMIS(True) 'Sending MMIS back to the beginning screen and checking for a password prompt
Call MMIS_case_number_finder(MMIS_case_number)

Call get_to_RKEY

'grabs the PMI number if one is listed on RKEY
If MMIS_case_number = "" Then
    EMReadscreen MMIS_case_number, 8, 9, 19
    MMIS_case_number= trim(MMIS_case_number)
End If

open_enrollment_case = FALSE
ask_about_oe = FALSE
nov_cut_off_date = #11/17/2021#
If Month(date) = 10 OR Month(date) = 11 Then
	If DateDiff("d", date, nov_cut_off_date) > 0 Then ask_about_oe = TRUE
End If

If ask_about_oe = TRUE Then
	ask_if_open_enrollment = MsgBox("Are you processing an Open Enrollment?", vbQuestion + vbYesNo, "Open Enrollment?")
	If ask_if_open_enrollment = vbYes Then
		enrollment_month = "01"
		enrollment_year = "22"
		open_enrollment_case = TRUE
		case_open_enrollment_yn = "Yes"
	End If
End If

IF open_enrollment_case = FALSE Then
	enrollment_month = CM_plus_1_mo
	enrollment_year = CM_plus_1_yr

	this_month = monthname(month(date))
	Select Case this_month
	    Case "January"
			cut_off_date = #01/20/2022#
	    Case "February"
			cut_off_date = #02/16/2022#
	    Case "March"
			cut_off_date = #03/22/2022#
	    Case "April"
			cut_off_date = #04/20/2022#
	    Case "May"
			cut_off_date = #05/19/2022#
	    Case "June"
			cut_off_date = #06/21/2022#
	    Case "July"
			cut_off_date = #07/20/2022#
	    Case "August"
			cut_off_date = #08/22/2022#
	    Case "September"
			cut_off_date = #09/21/2022#
	    Case "October"
			cut_off_date = #10/20/2022#
	    Case "November"
			cut_off_date = #11/17/2022#
	    Case "December"
			cut_off_date = #12/20/2022#
	End Select
	'MsgBox cut_off_date
	If cut_off_date <> "" Then
	    If DateDiff("d", date, cut_off_date) < 0 Then
	        'MsgBox DateDiff("d", date, cut_off_date)
	        enrollment_month = CM_plus_2_mo
	        enrollment_year = CM_plus_2_yr
	    End If
	End If
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 206, 180, "Enrollment Information"
  EditBox 90, 25, 60, 15, MMIS_case_number
  EditBox 90, 45, 20, 15, enrollment_month
  EditBox 115, 45, 20, 15, enrollment_year
  DropListBox 55, 75, 95, 15, "Select one..."+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Ucare"+chr(9)+"United Healthcare", Health_plan
  CheckBox 120, 95, 25, 10, "Yes", Insurance_yes
  CheckBox 120, 105, 25, 10, "Yes", foster_care_yes
  DropListBox 110, 120, 90, 45, "Select One..."+chr(9)+"Phone"+chr(9)+"Paper Enrollment Form"+chr(9)+"Morning Letters", enrollment_source
  DropListBox 110, 140, 50, 45, "No"+chr(9)+"Yes", case_open_enrollment_yn
  ButtonGroup ButtonPressed
    OkButton 95, 160, 50, 15
    CancelButton 150, 160, 50, 15
  GroupBox 5, 10, 150, 55, "Leading zeros not needed"
  Text 10, 30, 50, 10, "Case Number:"
  Text 10, 50, 80, 10, "Enrollment Month/Year:"
  Text 10, 80, 40, 10, "Health plan:"
  Text 10, 95, 100, 10, "Other Insurance for this case?"
  Text 10, 105, 50, 10, "Foster Care?"
  Text 10, 125, 100, 10, "Enrollment was requested via"
  Text 20, 145, 85, 10, "Is this Open Enrollment?"
EndDialog

'do the dialog here
Do
    err_msg = ""

	Dialog Dialog1
	cancel_without_confirmation

    MMIS_case_number = trim(MMIS_case_number)

    If MMIS_case_number = "" then err_msg = err_msg & vbNewLine & "* Enter the case number."
	If enrollment_source = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate where the request for the enrollment came from (phone call or enrollment form)."
	If case_open_enrollment_yn = "Yes" Then
		enrollment_month = "01"
		enrollment_year = "22"
		open_enrollment_case = TRUE
	Else
		If enrollment_month = "" OR enrollment_year = "" Then err_msg = err_msg & vbNewLine & "* Enter the month and year enrollment is effective."
	End If
    If health_plan = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select a health plan."

    If err_msg <> "" Then MsgBOx "Please resolve to continue: " & vbNewLine & err_msg
Loop until err_msg = ""
If case_open_enrollment_yn = "No" Then open_enrollment_case = FALSE
MAXIS_case_number = MMIS_case_number

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
Call get_to_RKEY

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
Call clear_line_of_text(4, 19)		'Clearing all of the search options used on RKEY as we must ONLY enter a case number
Call clear_line_of_text(5, 19)
Call clear_line_of_text(5, 48)
Call clear_line_of_text(6, 19)
Call clear_line_of_text(6, 48)
Call clear_line_of_text(6, 69)
Call clear_line_of_text(9, 19)
Call clear_line_of_text(9, 48)
Call clear_line_of_text(9, 69)

EMWriteScreen MMIS_case_number, 9, 19
transmit
'Now we are at RCAD
'We are going to grab the phone numbers while we are here.
phone_droplist = "Select or Type"
EMReadScreen phone_one, 12, 10, 30
phone_one = trim(phone_one)
phone_one = "(" & phone_one
If phone_one = "(-" Then phone_one = ""
phone_one = replace(phone_one, " ", ")")
If phone_one <> "" Then phone_droplist = phone_droplist+chr(9)+phone_one

EMReadScreen phone_two, 16, 11, 17
phone_two = trim(phone_two)
phone_two = "(" & phone_two
If phone_two = "(-" Then phone_two = ""
phone_two = replace(phone_two, " ", ")")
If phone_two <> "" Then phone_droplist = phone_droplist+chr(9)+phone_two

EMReadScreen phone_three, 16, 12, 17
phone_three = trim(phone_three)
phone_three = "(" & phone_three
If phone_three = "(-" Then phone_three = ""
phone_three = replace(phone_three, " ", ")")
If phone_three <> "" Then phone_droplist = phone_droplist+chr(9)+phone_three


'Now we continue to RCIN
transmit
transmit
EMReadscreen RCIN_check, 4, 1, 49
If RCIN_check <> "RCIN" then script_end_procedure_with_error_report("The listed Case number was not found. Check your Case number and try again.")

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

EMWriteScreen "X", 11, 2
transmit
EMWriteScreen "RCAS", 1, 8
transmit
rcas_row = 9
Do
	EMReadScreen rcas_case_numb, 8, rcas_row, 8
	' MsgBox rcas_case_numb
	If rcas_case_numb = MMIS_case_number Then
		EMReadScreen svc_loc, 3, rcas_row, 57
		EMReadScreen rcas_case_type, 1, rcas_row, 39
		' MsgBox "Servicing County - " & svc_loc & vbNewLine & "Type - " & rcas_case_type
		If svc_loc <> "027" Then script_end_procedure("It appears this case is either not serviced in Hennepin County. The script will now end as MMIS cannot be updated by a Hennepin County worker.")
		' If rcas_case_type <> "D" AND rcas_case_type <> "M" Then script_end_procedure("It appears this case is a case type not County Administered. The script will now end as MMIS cannot be updated by a Hennepin County worker.")
	End If
	rcas_row = rcas_row + 1
Loop until rcas_case_numb = "        "
PF3

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 250, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 105, 10, "Household members to look at:"
	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 175, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
	ButtonGroup ButtonPressed
	OkButton 195, 10, 50, 15
	CancelButton 195, 30, 50, 15
EndDialog

'runs the dialog that has been dynamically created. Streamlined with new functions.
Dialog Dialog1
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

const client_name  			= 0
const client_pmi   			= 1
const current_plan 			= 2
const current_plan_date		= 3
const new_plan     			= 4
const change_rsn   			= 5
const disenrol_rsn 			= 6
const med_code     			= 7
const dent_code    			= 8
const contr_code  			= 9
const preg_yes 	   			= 10
const interp_code  			= 11
const new_plan_two 			= 12
const contr_code_two 		= 13
const change_rsn_two 		= 14
const disenrol_rsn_two 		= 15
const first_new_plan 		= 16
const first_contr_code 		= 17
const first_change_rsn 		= 18
const first_disenrol_rsn 	= 19
const manual_enrollment 	= 20
const manual_enrollment_date	= 21
const manual_contr_code		= 22
const manual_new_plan		= 23
const manual_change_rsn		= 24
const manual_disenrol_rns	= 25
const enrol_sucs   			= 26

Dim MMIS_clients_array
ReDim MMIS_clients_array (enrol_sucs, 0)

EMReadScreen RCIN_check, 4, 1, 49
If RCIN_check = "RCIN" Then PF6
Call get_to_RKEY

item = 0

For each member in HH_member_array
	ReDim Preserve MMIS_clients_array(enrol_sucs, item)
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
	MMIS_clients_array(manual_enrollment, item) = FALSE

	'check RPOL to see if there is other insurance available, if so worker processes manually
	'EMWriteScreen "X", 11, 2
	'Transmit
	EMWriteScreen "RPOL", 1, 8
	transmit
	'making sure script got to right panel
	EMReadScreen RPOL_check, 4, 1, 52
	If RPOL_check <> "RPOL" then script_end_procedure_with_error_report("The script was unable to navigate to RPOL process manually if needed.")

	EMreadscreen policy_number, 1, 7, 8
    If policy_number <> " " then

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 161, 145, "RPOL Updated"
          CheckBox 20, 100, 125, 10, "RPOL ended/ ready for enrollment", rpol_ended_checkbox
          ButtonGroup ButtonPressed
            OkButton 105, 125, 50, 15
          Text 10, 10, 145, 25, "The script has found information on RPOL. The script cannot review RPOL to determine if enrollment information can be changed. "
          GroupBox 10, 45, 145, 70, "REVIEW RPOL"
          Text 50, 60, 65, 10, "*** Check RPOL ***"
          Text 30, 75, 105, 15, "Review to see if the enrollment being attempted can be added."
        EndDialog

        dialog Dialog1

        If rpol_ended_checkbox = unchecked Then
            PF6
    		script_end_procedure_with_error_report ("This case has spans on RPOL. Please evaluate manually at this time.")
        End If
    End If

	PF6

	EMWriteScreen "RPPH", 1, 8
	transmit
	row = 1
	col = 1
	EMSearch "99/99/99", row, col
	IF row < 10 Then
		If col = 18 Then
			EMReadScreen excl_code, 2, row, 2
			EMReadScreen hp_curr_start_date, 8, row, 9
		ElseIf col = 45 Then
			EMReadScreen excl_code, 2, row, 29
			EMReadScreen hp_curr_start_date, 8, row, 36
		ElseIf col = 72 Then
			EMReadScreen excl_code, 2, row, 56
			EMReadScreen hp_curr_start_date, 8, row, 63
		End If
		MMIS_clients_array(current_plan_date, item) = hp_curr_start_date

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
		If hp_code = "A168407400" then MMIS_clients_array(current_plan, item) = "United Healthcare"
		If hp_code = "A836618200" then MMIS_clients_array(current_plan, item) = "Hennepin Health PMAP"
		If hp_code = "A965713400" then MMIS_clients_array(current_plan, item) = "Hennepin Health SNBC"
		EMReadScreen hp_curr_start_date, 8, row, 5
		MMIS_clients_array(current_plan_date, item) = hp_curr_start_date
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
dlg_len = 60
If enrollment_source = "Phone" OR enrollment_source = "Paper Enrollment Form" Then
    dlg_len = dlg_len + 20
End If

name_list = ""
For person = 0 to Ubound(MMIS_clients_array, 2)
    name_list = name_list & +chr(9)+MMIS_clients_array(first_name_ini, person)
    dlg_len = dlg_len + 20
Next

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 750, dlg_len, "Enrollment Information"
  Text 5, 5, 25, 10, "Name"
  Text 100, 5, 15, 10, "PMI"
  Text 143, 5, 100, 10, "Current Plan/Exclusion and the"
  Text 240, 5, 45, 10, " Start Date"
  Text 285, 5, 45, 10, "Med Clinic"
  Text 340, 5, 45, 10, "Den Clinic"
  Text 390, 5, 40, 10, "Health plan:"
  Text 460, 5, 55, 10, "Contract Code:"
  Text 520, 5, 55, 10, "Change reason:"
  Text 585, 5, 60, 10, "Disenroll reason:"
  Text 650, 5, 35, 10, "Pregnant?"
  Text 690, 5, 55, 10, "Interpreter Code"

  For person = 0 to Ubound(MMIS_clients_array, 2)
    If enrollment_source = "Morning Letters" Then MMIS_clients_array(change_rsn, person) = "Reenrollment"
	If open_enrollment_case = TRUE Then
		MMIS_clients_array(change_rsn, person) = "Open enrollment"
		MMIS_clients_array(disenrol_rsn, person) = "Open Enrollment"
	End If
  	Text 5, (x * 20) + 25, 95, 10, MMIS_clients_array(client_name, person)
  	Text 100, (x * 20) + 25, 35, 10, MMIS_clients_array(client_pmi, person)
  	Text 143, (x * 20) + 25, 95, 10, MMIS_clients_array(current_plan, person)
	Text 240, (x * 20) + 25, 35, 10, MMIS_clients_array(current_plan_date, person)
  	EditBox 285, (x * 20) + 20, 45, 15, MMIS_clients_array(med_code, person)
  	EditBox 340, (x * 20) + 20, 45, 15, MMIS_clients_array(dent_code, person)
    DropListBox 390, (x * 20) + 20, 60, 15, " "+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Ucare"+chr(9)+"United Healthcare", MMIS_clients_array(new_plan, person)
  	DropListBox 460, (x * 20) + 20, 50, 15, "MA 12"+chr(9)+"NM 12"+chr(9)+"MA 30"+chr(9)+"MA 35"+chr(9)+"IM 12", MMIS_clients_array(contr_code, person)
	DropListBox 520, (x * 20) + 20, 60, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", MMIS_clients_array(change_rsn, person)
  	DropListBox 585, (x * 20) + 20, 60, 15, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary"+chr(9)+"DELETE SPAN", MMIS_clients_array(disenrol_rsn, person)
  	CheckBox 655, (x * 20) + 20, 25, 10, "Yes", MMIS_clients_array(preg_yes, person)
	EditBox 700, (x * 20) + 20, 25, 15, MMIS_clients_array(interp_code, person)
	x = x + 1
  Next

  Text 5, (x * 20) + 25, 45, 10, "Other Notes:"
  EditBox 55, (x * 20) + 20, 690, 15, other_notes

  If enrollment_source = "Phone" Then
      GroupBox 5, (x * 20) + 40, 530, 35, "Phone Call Information"
      Text 10, (x * 20) + 60, 40, 10, "Caller name"
      ComboBox 55, (x * 20) + 55, 120, 45, " " & name_list, caller_name
      Text 175, (x * 20) + 60, 40, 10, ", who is the"
      ComboBox 215, (x * 20) + 55, 80, 45, "Client"+chr(9)+"AREP", caller_rela
      CheckBox 305, (x * 20) + 50, 30, 10, "Used", used_interpreter_checkbox
	  Text 305, (x * 20) + 60, 70, 10, "Interpreter"
	  Text 350, (x * 20) + 60, 80, 10, "Phone Number of Caller"
	  ComboBox 430, (x * 20) + 55, 100, 15, phone_droplist, phone_number_of_caller
      x = x + 1
  End If
  If enrollment_source = "Paper Enrollment Form" Then
	  GroupBox 5, (x * 20) + 45, 180, 30, "Paper Form Information"
	  Text 10, (x * 20) + 60, 80, 10, "Form Received Date:"
	  EditBox 95, (x * 20) + 55, 80, 15, form_received_date
  End If

  If enrollment_source = "Paper Enrollment Form" OR enrollment_source = "Phone" Then
	  Text 570, dlg_len - 35, 60, 10, "Worker Signature"
	  EditBox 635, dlg_len - 40, 110, 15, worker_signature
  Else
	  Text 445, dlg_len - 15, 60, 10, "Worker Signature"
	  EditBox 510, dlg_len - 20, 110, 15, worker_signature
  End If
  ButtonGroup ButtonPressed
    OkButton 640, dlg_len - 20, 50, 15
    CancelButton 695, dlg_len - 20, 50, 15
EndDialog

Do
    err_msg = ""

	Dialog Dialog1
	cancel_confirmation

    For person = 0 to Ubound(MMIS_clients_array, 2)
        If left(MMIS_clients_array(current_plan, person), 3) <> "XCL" AND trim(MMIS_clients_array(current_plan, person)) <> "" Then
            If MMIS_clients_array(disenrol_rsn, person) = "Select one..." Then err_msg = err_msg & vbNewLine & "* Since " & MMIS_clients_array(client_name, person) & " is currently on a health plan, please select a disenrollment reason for the " & MMIS_clients_array(current_plan, person) & " plan."
        End If
        If MMIS_clients_array(change_rsn, person) = "Select one..." Then err_msg = err_msg & vbNewLine & "* Select a reason to enroll  " & MMIS_clients_array(client_name, person) & " into a new plan."
    Next

    If enrollment_source = "Phone" Then

        If trim(caller_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the caller."
        If trim(caller_rela) = "" Then err_msg = err_msg & vbNewLine & "* Select who is calling (typically Client or AREP)."
		If trim(phone_number_of_caller) = "" Then err_msg = err_msg & vbNewLine & "* Enter the phone number the person is calling from."

    End If

    If worker_signature = "" THen err_msg = err_msg & vbNewLine & "* Enter your name for the case note signature."

    If err_msg <> "" THen MsgBox "Please resovle to continue:" & vbNewLine & err_msg
	If ButtonPressed = 0 Then err_msg = "LOOP"

Loop Until err_msg = ""

process_manually_message = ""
original_enrollment_date = enrollment_date

script_run_lowdown = script_run_lowdown & "Enrollment source - " & enrollment_source & vbCr
If enrollment_source = "Phone" Then script_run_lowdown = script_run_lowdown & "Caller name: " & caller_name & " - Relationship: " & caller_rela  & vbCr
If enrollment_source = "Paper Enrollment Form" Then script_run_lowdown = script_run_lowdown & "Form Received Date: " & form_received_date & vbCr

For person = 0 to Ubound(MMIS_clients_array, 2)
	script_run_lowdown = script_run_lowdown & "Name: " & MMIS_clients_array(client_name, person) & " - PMI: " & MMIS_clients_array(client_pmi, person) & vbCr
	script_run_lowdown = script_run_lowdown & "Current Plan: " & MMIS_clients_array(current_plan, person) & vbCr
	script_run_lowdown = script_run_lowdown & "New Plan: " & MMIS_clients_array(new_plan, person) & " - Contract Code: " & MMIS_clients_array(contr_code, person) & vbCr
	script_run_lowdown = script_run_lowdown & "Change Reason: " & MMIS_clients_array(change_rsn, person) & vbCr
	script_run_lowdown = script_run_lowdown & "Disenrollment Reason: " & MMIS_clients_array(disenrol_rsn, person) & vbCr
	script_run_lowdown = script_run_lowdown & "---------------------------------------------" & vbCr & vbCr
Next

If MNSURE_Case = TRUE Then
	For member = 0 to Ubound(MMIS_clients_array, 2)
		first_attempt = TRUE
		updated_manually = FALSE
		replace_enrollment_span = FALSE
		enrollment_date = original_enrollment_date
		MMIS_clients_array(first_new_plan, member) = MMIS_clients_array(new_plan, member)
		MMIS_clients_array(first_contr_code, member) = MMIS_clients_array(contr_code, member)
		MMIS_clients_array(first_change_rsn, member) = MMIS_clients_array(change_rsn, member)
		MMIS_clients_array(first_disenrol_rsn, member) = MMIS_clients_array(disenrol_rsn, member)
		Do
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
			If MMIS_clients_array(disenrol_rsn, member) = "DELETE SPAN" Then
				disenrollment_reason = ""
				replace_enrollment_span = TRUE
			End If

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

			If first_attempt = TRUE Then
				Call get_to_RKEY

				'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
				EMWriteScreen "c", 2, 19
				EMWriteScreen "        ", 9, 19
				EMWriteScreen MMIS_clients_array(client_pmi, member), 4, 19
				transmit
				EMReadscreen RKEY_check, 4, 1, 52
				If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine

				rpol_attempt = 1
				Do
					'check RPOL to see if there is other insurance available, if so worker processes manually
					EMWriteScreen "rpol", 1, 8
					transmit
					'making sure script got to right panel
					EMReadScreen RPOL_check, 4, 1, 52
					rpol_attempt = rpol_attempt + 1
					If rpol_attempt = 20 Then exit do
				Loop until RPOL_check = "RPOL"

				If RPOL_check <> "RPOL" then process_manually_message = process_manually_message & "Could not navigate to RPOL for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
				EMreadscreen policy_number, 1, 7, 8
	            If policy_number <> " " then

	                Dialog1 = ""
	                BeginDialog Dialog1, 0, 0, 161, 145, "RPOL Updated"
	                  CheckBox 20, 100, 125, 10, "RPOL ended/ ready for enrollment", rpol_ended_checkbox
	                  ButtonGroup ButtonPressed
	                    OkButton 105, 125, 50, 15
	                  Text 10, 10, 145, 25, "The script has found information on RPOL. The script cannot review RPOL to determine if enrollment information can be changed. "
	                  GroupBox 10, 45, 145, 70, "REVIEW RPOL"
	                  Text 50, 60, 65, 10, "*** Check RPOL ***"
	                  Text 30, 75, 105, 15, "Review to see if the enrollment being attempted can be added."
	                EndDialog

	                dialog Dialog1

	                If rpol_ended_checkbox = unchecked Then process_manually_message = process_manually_message & "RPOL for PMI " & MMIS_clients_array(client_pmi, member) & " has a span listed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
	            End If
				'nav to RPPH
				EMWriteScreen "rpph", 1, 8
				transmit

				'making sure script got to right panel
				EMReadScreen RPPH_check, 4, 1, 52
				If RPPH_check <> "RPPH" then process_manually_message = process_manually_message & "Could not navigate to RPPH for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				first_attempt = FALSE
			End If

			xcl_end_date = DateAdd("d", -1, enrollment_date)
			xcl_end_month = right("00" & DatePart("m", xcl_end_date), 2)
			xcl_end_day   = right("00" & DatePart("d", xcl_end_date), 2)
			xcl_end_year  = right(DatePart("yyyy", xcl_end_date), 2)
			xcl_end_date  = xcl_end_month & "/" & xcl_end_day & "/" & xcl_end_year

			enrl_month = right("00" & DatePart("m", enrollment_date), 2)
			enrl_day   = right("00" & DatePart("d", enrollment_date), 2)
			enrl_year  = right(DatePart("yyyy", enrollment_date), 2)
			enrollment_date  = enrl_month & "/" & enrl_day & "/" & enrl_year
			' msgbox enrollment_date & vbNewLine & xcl_end_date
			'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
			If left(MMIS_clients_array(current_plan, member), 3) = "XCL" Then
				If MMIS_clients_array(current_plan, member) = "XCL - Delayed Decision" OR MMIS_clients_array(current_plan, member) = "XCL - Private HMO Coverage" OR MMIS_clients_array(current_plan, member) = "XCL - Adoption Assistance" Then
					row = 1
					col = 1
					EMSearch "99/99/99", row, col
                    If col <> 0 Then
                        EMReadScreen beg_of_excl, 8, row, 9
                        IF beg_of_excl = enrollment_date Then
                            EMWriteScreen "...", row, 2
                        Else
                            EMWriteScreen xcl_end_date, row, col
                        End if
                    End If
				Else
					process_manually_message = process_manually_message & "There is an exclusion code other than 'YY' for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				End If
			End If
			EMReadscreen XCL_code, 2, 6, 2
			If XCL_code = "* " then
				EMSetCursor 6, 2
				EMSendKey "..."
			End if
			' msgbox "Exclusion ended"

			If MMIS_clients_array(new_plan, member) = "Health Partners" then health_plan_code = "A585713900"
			If MMIS_clients_array(new_plan, member) = "Ucare" then health_plan_code = "A565813600"
			If MMIS_clients_array(new_plan, member) = "Medica" then health_plan_code = "A405713900"
			If MMIS_clients_array(new_plan, member) = "Blue Plus" then health_plan_code = "A065813800"
			If MMIS_clients_array(new_plan, member) = "United Healthcare" then health_plan_code = "A168407400"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health PMAP" then health_plan_code = "A836618200"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health SNBC" then health_plan_code = "A965713400"

			contract_code = MMIS_clients_array (contr_code, member)
			Contract_code_part_one = left(contract_code, 2)
			Contract_code_part_two = right(contract_code, 2)

            If process_manually_message = "" Then
    			'enter disenrollment reason
				If replace_enrollment_span = TRUE Then
					EMWriteScreen "...", 13, 5
					transmit
					call write_value_and_transmit("RPPH", 1, 8)
				End If
                If disenrollment_reason <> "" Then
                    EMReadScreen beg_of_curr_span, 8, 13, 5
                    If beg_of_curr_span = enrollment_date Then
                        EMWriteScreen "...", 13, 5
                    Else
                        EMWriteScreen xcl_end_date, 13, 14
                        EMWriteScreen disenrollment_reason, 13, 75
                    End If
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
				' MsgBox "STOP HERE"
    			EMWaitReady 0, 0

    			EMReadScreen false_end, 8, 14, 14
    			If false_end = "99/99/99" Then
    				EMReadScreen double_check, 2, 14, 5
    				If double_check = "  " Then EMWriteScreen "...", 14, 5
    			End If

				transmit
				EMReadScreen current_panel, 4, 1, 52
    			EMReadScreen RPPH_error_check, 78, 24, 2

				RPPH_error_check = trim(RPPH_error_check)
				enrollment_successful = TRUE

				' MsgBox "Panel - " & current_panel & vbCr & "Error - " & RPPH_error_check
				If current_panel <> "REFM" Then
					enrollment_successful = FALSE
				' ElseIf RPPH_error_check <> "" Then

				Else
					'updating REFM'
					call enter_detail_on_refm

					'nav to RPPH
					EMWriteScreen "rpph", 1, 8
					transmit

					EMReadScreen panel_enrollment_date, 8, 13, 5
					'enter managed care plan code
					EMReadScreen panel_health_plan_code, 10, 13, 23
					'enter contract code
					EMReadScreen panel_contract_code_part_one, 2, 13, 34
					EMReadScreen panel_contract_code_part_two, 2, 13, 37
					'enter change reason
					EMReadScreen panel_change_reason, 2, 13, 71

					If panel_enrollment_date <> enrollment_date Then enrollment_successful = FALSE
					If panel_health_plan_code <> health_plan_code Then enrollment_successful = FALSE
					If panel_contract_code_part_one <> contract_code_part_one Then enrollment_successful = FALSE
					If panel_contract_code_part_two <> contract_code_part_two Then enrollment_successful = FALSE
					If panel_change_reason <> change_reason Then enrollment_successful = FALSE
					' MsgBox enrollment_successful
					If enrollment_successful = TRUE Then
						EMWriteScreen "refm", 1, 8
		                transmit
					End If
				End If

				If enrollment_successful = FALSE Then
					testing_run = TRUE
					If new_enrol_date = "" Then new_enrol_date = enrollment_date
					new_enrol_date = new_enrol_date & ""
					script_run_lowdown = script_run_lowdown & "Enrollment was not successful for " & MMIS_clients_array(client_name, member) & vbCr

					If MMIS_clients_array(new_plan_two, member) = "" Then MMIS_clients_array(new_plan_two, member) = MMIS_clients_array(new_plan, member)
					If MMIS_clients_array(contr_code_two, member) = "" Then MMIS_clients_array(contr_code_two, member) = MMIS_clients_array(contr_code, member)
					If MMIS_clients_array(change_rsn_two, member) = "" Then MMIS_clients_array(change_rsn_two, member) = MMIS_clients_array(change_rsn, member)
					If MMIS_clients_array(disenrol_rsn_two, member) = "" Then MMIS_clients_array(disenrol_rsn_two, member) = MMIS_clients_array(disenrol_rsn, member)

					Do
						err_msg = ""

						BeginDialog Dialog1, 0, 0, 456, 290, "Update Enrollment Options due to MMIS Failure or Error"
						  DropListBox 80, 185, 125, 45, "Select One..."+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Ucare"+chr(9)+"United Healthcare", MMIS_clients_array(new_plan_two, member)
						  DropListBox 250, 185, 65, 45, "Select One..."+chr(9)+"MA 12"+chr(9)+"NM 12"+chr(9)+"MA 30"+chr(9)+"MA 35"+chr(9)+"IM 12", MMIS_clients_array (contr_code_two, member)
						  EditBox 390, 185, 50, 15, new_enrol_date
						  DropListBox 80, 205, 130, 45, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", MMIS_clients_array(change_rsn_two, member)
						  DropListBox 100, 225, 130, 45, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", MMIS_clients_array(disenrol_rsn_two, member)
						  ButtonGroup ButtonPressed
						    PushButton 10, 270, 135, 15, "Skip Enrollment for " & MMIS_clients_array(client_name, member), skip_enrollment_btn
						    PushButton 160, 270, 150, 15, "Enrollment Already Updated Manually", read_manual_enrollment_btn
						    PushButton 325, 270, 125, 15, " Retry Enrollment with Script", retry_enrollment_btn
						  Text 10, 10, 345, 10, "The selections made for this person caused a message or error in MMIS and were not able to be updated."
						  GroupBox 10, 25, 440, 105, "Initial Selections for " & MMIS_clients_array(client_name, member)
						  Text 20, 45, 190, 10, "Plan to Enroll:" & MMIS_clients_array(new_plan, member)
						  Text 215, 45, 75, 10, "ID / Desc:" & MMIS_clients_array (contr_code, member)
						  Text 330, 45, 110, 10, "Enrollment Date:" & enrollment_date
						  Text 20, 60, 210, 10, "Change Reason:" & MMIS_clients_array(change_rsn, member)
						  Text 20, 75, 220, 10, "Disenrollment Reason:" & MMIS_clients_array(disenrol_rsn, member)
						  Text 20, 90, 55, 10, "Pregnant: " & pregnant_yn
						  Text 115, 90, 60, 10, "Interpreter: " & interpreter_yn
						  Text 190, 90, 90, 10, "Interpreter Code:" & MMIS_clients_array(interp_code, member)
						  Text 20, 115, 215, 10, "Current Enrollment:" & MMIS_clients_array(current_plan, member)
						  GroupBox 10, 135, 440, 30, "MMIS Error Messages"
						  Text 20, 150, 415, 10, RPPH_error_check
						  GroupBox 10, 170, 440, 95, "Change Selections for " & MMIS_clients_array(client_name, member)
						  Text 20, 190, 60, 10, "Enrollment Plan:"
						  Text 215, 190, 35, 10, "ID / Desc:"
						  Text 330, 190, 60, 10, "Enrollment Date:"
						  Text 20, 210, 60, 10, "Change Reason: "
						  Text 20, 230, 80, 10, "Disenrollment Reason:"
						EndDialog

						dialog Dialog1
						cancel_confirmation

						If ButtonPressed = retry_enrollment_btn Then

							If MMIS_clients_array(new_plan_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "*  Select which plan you need to enroll " & MMIS_clients_array(client_name, person) & "into."
							If MMIS_clients_array (contr_code_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the product ID and description for enrollment."
							If MMIS_clients_array(change_rsn_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select a reason to enroll  " & MMIS_clients_array(client_name, person) & " into a new plan."
							If left(MMIS_clients_array(current_plan, person), 3) <> "XCL" AND trim(MMIS_clients_array(current_plan, person)) <> "" Then
								If MMIS_clients_array(disenrol_rsn_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since " & MMIS_clients_array(client_name, person) & " is currently on a health plan, please select a disenrollment reason for the " & MMIS_clients_array(current_plan, person) & " plan."
							End If
							If err_msg <> "" THen MsgBox "Please resovle to continue:" & vbNewLine & err_msg

						End If
					Loop until err_msg = ""

					new_enrol_date = DateAdd("d", 0, new_enrol_date)
					If ButtonPressed = retry_enrollment_btn Then
						PF10
						MMIS_clients_array(new_plan, member) = MMIS_clients_array(new_plan_two, member)
						MMIS_clients_array(contr_code, member) = MMIS_clients_array(contr_code_two, member)
						MMIS_clients_array(change_rsn, member) = MMIS_clients_array(change_rsn_two, member)
						MMIS_clients_array(disenrol_rsn, member) = MMIS_clients_array(disenrol_rsn_two, member)

						enrollment_date = new_enrol_date
						script_run_lowdown = script_run_lowdown & "Button Pressed - RETRY ENROLLMENT" & vbCr
					End If
					If ButtonPressed = read_manual_enrollment_btn Then
						EMReadScreen current_panel, 4, 1, 52

						If current_panel = "RPPH" Then
							Call read_detail_on_RPPH(MMIS_clients_array(manual_enrollment_date, member), panel_health_plan_code, panel_contract_code_part_one, panel_contract_code_part_two, MMIS_clients_array(manual_contr_code, member), panel_change_reason, panel_disenrollment_reason, MMIS_clients_array(manual_new_plan, member), MMIS_clients_array(manual_change_rsn, member), MMIS_clients_array(manual_disenrol_rns, member))

							EMWriteScreen "refm", 1, 8
			                transmit

							call enter_detail_on_refm
						ElseIf current_panel = "REFM" Then
							call enter_detail_on_refm

							EMWriteScreen "rpph", 1, 8
							transmit

							Call read_detail_on_RPPH(MMIS_clients_array(manual_enrollment_date, member), panel_health_plan_code, panel_contract_code_part_one, panel_contract_code_part_two, MMIS_clients_array(manual_contr_code, member), panel_change_reason, panel_disenrollment_reason, MMIS_clients_array(manual_new_plan, member), MMIS_clients_array(manual_change_rsn, member), MMIS_clients_array(manual_disenrol_rns, member))
						End If
						enrollment_successful = TRUE
						MMIS_clients_array(manual_enrollment, member) = TRUE
						PF9
						Do
							PF3
							EMReadScreen where_are_we, 4, 1, 52
						Loop until where_are_we = "RKEY"
						script_run_lowdown = script_run_lowdown & "Button Pressed - MANUAL ENROLLMENT" & vbCr
					End If
					If ButtonPressed = skip_enrollment_btn Then
						process_manually_message = process_manually_message & "Enrollment was cancelled by you when you pressed the button 'Skip Enrollment for " & MMIS_clients_array(client_name, member) & "' button on the 'Update Enrollment Options due to MMIS Failure or Error' dialog." & vbNewLine & vbNewLine & ""
						enrollment_successful = TRUE
						PF10
						script_run_lowdown = script_run_lowdown & "Button Pressed - SKIP ENROLLMENT" & vbCr
					End If

					first_attempt = FALSE
				End If
            ELSE
                'REFM screen
                EMWriteScreen "refm", 1, 8
                transmit
            End If

	        'blanking out varibles if the other option is selected
	        If change_reason = "Select one..." then change_reason = ""
	        If disenrollment_reason = "Select one..." then disenrollment_reason = ""
			'making sure script got to right panel
			' EMReadScreen REFM_check, 4, 1, 52
			' If REFM_check <> "REFM" then process_manually_message = process_manually_message & "The script was unable to navigate to REFM for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
		Loop until enrollment_successful = TRUE

        ' If enrollment_source = "Paper Enrollment Form" Then
    	' 	'form rec'd
    	' 	EMsetcursor 10, 16
    	' 	EMSendkey "y"
    	' 	'other insurance y/n
    	' 	EMsetcursor 11, 18
    	' 	EMsendkey insurance_yn
    	' 	'preg y/n
    	' 	EMsetcursor 12, 19
    	' 	EMsendkey pregnant_yn
    	' 	'interpreter y/n
    	' 	EMsetcursor 13, 29
    	' 	EMsendkey interpreter_yn
    	' 	'interpreter type
    	' 	if MMIS_clients_array(interp_code, member) <> "" then
    	' 		EMsetcursor 13, 52
    	' 		EMsendKey MMIS_clients_array(interp_code, member)
    	' 	end if
    	' 	'medical clinic code
    	' 	EMsetcursor 19, 4
    	' 	EMsendkey MMIS_clients_array(med_code, member)
    	' 	'dental clinic code if applicable
    	' 	EMsetcursor 19, 24
    	' 	EMsendkey MMIS_clients_array(dent_code, member)
    	' 	'foster care y/n
    	' 	EMsetcursor 21, 15
    	' 	EMsendkey foster_care_yn
		'     ' msgbox "REFM updated"
        ' Else
        '     'form rec'd
        '     EMsetcursor 10, 16
        '     EMSendkey "n"
        ' End If
		'
		' PF9
		'
		' 'error handling to ensure that enrollment date and exclusion dates don't conflict
		' EMReadScreen REFM_error_check, 19, 24, 2 'checks for an inhibiting edit
		' IF REFM_error_check <> "                   " then
        '     IF REFM_error_check <> "INVALID KEY ENTERED" AND REFM_error_check <> "INVALID KEY PRESSED" then
        '         EMReadScreen full_error_msg, 79, 24, 2
        '         full_error_msg = trim(full_error_msg)
		' 	    process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & full_error_msg & vbNewLine & vbNewLine
		'     END IF
        ' END IF


		If MMIS_clients_array(manual_enrollment, member) <> TRUE Then
			PF9
			If MMIS_clients_array(current_plan, member) = "XCL - Adoption Assistance" Then PF9

			'error handling to ensure that enrollment date and exclusion dates don't conflict
			EMReadScreen REFM_error_check, 19, 24, 2 'checks for an inhibiting edit
			IF REFM_error_check <> "                   " then
	            IF REFM_error_check <> "INVALID KEY ENTERED" AND REFM_error_check <> "INVALID KEY PRESSED" then
	                EMReadScreen full_error_msg, 79, 24, 2
	                full_error_msg = trim(full_error_msg)
				    process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & full_error_msg & vbNewLine & vbNewLine
			    END IF
	        END IF
		End If
		' msgbox "all updated - see casenote code"
		'Save and case note
		EMReadScreen where_are_we, 4, 1, 52
		Do While where_are_we <> "RKEY"
			pf3
			EMReadScreen where_are_we, 4, 1, 52
		Loop

		EMWriteScreen "i", 2, 19
        EMWriteScreen "        ", 9, 19
        EMWriteScreen MMIS_clients_array(client_pmi, member), 4, 19
		transmit
		check_enrollment = TRUE
		IF ButtonPressed = read_manual_enrollment_btn Then check_enrollment = FALSE
		IF ButtonPressed = skip_enrollment_btn Then check_enrollment = FALSE
		EMReadScreen rsum_enrollment, 8, 16, 20
		EMReadScreen rsum_plan, 10, 16, 52
		' MsgBox "RSUM date and plan: " & rsum_enrollment & " - " & rsum_plan & vbNewLine & "Coded date and plan: " & Enrollment_date & " - " & health_plan_code
		IF rsum_enrollment = Enrollment_date AND rsum_plan = health_plan_code Then
			MMIS_clients_array(enrol_sucs, member) = TRUE
			script_run_lowdown = script_run_lowdown & "Enrollment appears successful for " & MMIS_clients_array(client_name, member) & vbCr
		''			pf4
		''			pf11
		''			EMSendkey "***HMO Note*** " &  MMIS_clients_array(client_name, member) & " enrolled into " &  MMIS_clients_array(new_plan, member) & " " & Enrollment_date & " " & worker_signature
		''			pf3
		Else
			failed_enrollment_message = failed_enrollment_message & vbNewLine & vbNewLine & process_manually_message
			MMIS_clients_array(enrol_sucs, member) = FALSE
			script_run_lowdown = script_run_lowdown & "Enrollment appears to have failed for " & MMIS_clients_array(client_name, member) & vbCr & "Process manually message:" & vbCr & process_manually_message & "-----------------" & vbCr
		End If
		' MsgBox process_manually_message
		pf3
		IF REFM_error_check = "WARNING: MA12,01/16" Then
			PF3
		END IF
		process_manually_message = ""

	Next
Else
	For member = 0 to Ubound(MMIS_clients_array, 2)
		first_attempt = TRUE
		updated_manually = FALSE
		replace_enrollment_span = FALSE
		enrollment_date = original_enrollment_date
		MMIS_clients_array(first_new_plan, member) = MMIS_clients_array(new_plan, member)
		MMIS_clients_array(first_contr_code, member) = MMIS_clients_array(contr_code, member)
		MMIS_clients_array(first_change_rsn, member) = MMIS_clients_array(change_rsn, member)
		MMIS_clients_array(first_disenrol_rsn, member) = MMIS_clients_array(disenrol_rsn, member)
		Do
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
			If MMIS_clients_array(disenrol_rsn, member) = "DELETE SPAN" Then
				disenrollment_reason = ""
				replace_enrollment_span = TRUE
			End If

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

			If first_attempt = TRUE Then
				Call get_to_RKEY
				' msgbox "At RKEY"
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
				' msgbox "person selected"
				transmit
				' msgbox "at RSUM"
				EMReadscreen RKEY_check, 4, 1, 52
				If RKEY_check = "RKEY" then process_manually_message = process_manually_message & "PMI " & MMIS_clients_array(client_pmi, member) & " could not be accessed. The enrollment for " & MMIS_clients_array(client_name, member) & " needs to be processed manually." & vbNewLine & vbNewLine
				' msgbox process_manually_message

				rpol_attempt = 1
				Do
					'check RPOL to see if there is other insurance available, if so worker processes manually
					EMWriteScreen "rpol", 1, 8
					transmit
					'making sure script got to right panel
					EMReadScreen RPOL_check, 4, 1, 52
					rpol_attempt = rpol_attempt + 1
					If rpol_attempt = 20 Then exit do
				Loop until RPOL_check = "RPOL"

				If RPOL_check <> "RPOL" then process_manually_message = process_manually_message & "Could not navigate to RPOL for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				EMreadscreen policy_number, 1, 7, 8
				If policy_number <> " " then

	                Dialog1 = ""
	                BeginDialog Dialog1, 0, 0, 161, 145, "RPOL Updated"
	                  CheckBox 20, 100, 125, 10, "RPOL ended/ ready for enrollment", rpol_ended_checkbox
	                  ButtonGroup ButtonPressed
	                    OkButton 105, 125, 50, 15
	                  Text 10, 10, 145, 25, "The script has found information on RPOL. The script cannot review RPOL to determine if enrollment information can be changed. "
	                  GroupBox 10, 45, 145, 70, "REVIEW RPOL"
	                  Text 50, 60, 65, 10, "*** Check RPOL ***"
	                  Text 30, 75, 105, 15, "Review to see if the enrollment being attempted can be added."
	                EndDialog

	                dialog Dialog1

	                If rpol_ended_checkbox = unchecked Then process_manually_message = process_manually_message & "RPOL for PMI " & MMIS_clients_array(client_pmi, member) & " has a span listed. The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
	            End If
				'nav to RPPH
				EMWriteScreen "rpph", 1, 8
				transmit

				'making sure script got to right panel
				EMReadScreen RPPH_check, 4, 1, 52
				If RPPH_check <> "RPPH" then process_manually_message = process_manually_message & "Could not navigate to RPPH for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				first_attempt = FALSE
			End If

			xcl_end_date = DateAdd("d", -1, enrollment_date)
			xcl_end_month = right("00" & DatePart("m", xcl_end_date), 2)
			xcl_end_day   = right("00" & DatePart("d", xcl_end_date), 2)
			xcl_end_year  = right(DatePart("yyyy", xcl_end_date), 2)
			xcl_end_date  = xcl_end_month & "/" & xcl_end_day & "/" & xcl_end_year

			enrl_month = right("00" & DatePart("m", enrollment_date), 2)
			enrl_day   = right("00" & DatePart("d", enrollment_date), 2)
			enrl_year  = right(DatePart("yyyy", enrollment_date), 2)
			enrollment_date  = enrl_month & "/" & enrl_day & "/" & enrl_year
			' msgbox enrollment_date & vbNewLine & xcl_end_date
			'Checks for exclusion code only deletes if YY or blank, if any other span entered it stops script.
			If left(MMIS_clients_array(current_plan, member), 3) = "XCL" Then
				If MMIS_clients_array(current_plan, member) = "XCL - Delayed Decision" OR MMIS_clients_array(current_plan, member) = "XCL - Private HMO Coverage" OR MMIS_clients_array(current_plan, member) = "XCL - Adoption Assistance" Then
					row = 1
					col = 1
					EMSearch "99/99/99", row, col
                    If col <> 0 Then
                        EMReadScreen beg_of_excl, 8, row, 9
                        IF beg_of_excl = enrollment_date Then
                            EMWriteScreen "...", row, 2
                        Else
                            EMWriteScreen xcl_end_date, row, col
                        End if
                    End If
				Else
					process_manually_message = process_manually_message & "There is an exclusion code other than 'YY' for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
				End If
			End If
			EMReadscreen XCL_code, 2, 6, 2
			If XCL_code = "* " then
				EMSetCursor 6, 2
				EMSendKey "..."
			End if
			' msgbox "Exclusion ended"

			If MMIS_clients_array(new_plan, member) = "Health Partners" then health_plan_code = "A585713900"
			If MMIS_clients_array(new_plan, member) = "Ucare" then health_plan_code = "A565813600"
			If MMIS_clients_array(new_plan, member) = "Medica" then health_plan_code = "A405713900"
			If MMIS_clients_array(new_plan, member) = "Blue Plus" then health_plan_code = "A065813800"
			If MMIS_clients_array(new_plan, member) = "United Healthcare" then health_plan_code = "A168407400"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health PMAP" then health_plan_code = "A836618200"
			If MMIS_clients_array(new_plan, member) = "Hennepin Health SNBC" then health_plan_code = "A965713400"

			contract_code = MMIS_clients_array (contr_code, member)
			Contract_code_part_one = left(contract_code, 2)
			Contract_code_part_two = right(contract_code, 2)

            If process_manually_message = "" Then
    			'enter disenrollment reason
				If replace_enrollment_span = TRUE Then
					EMWriteScreen "...", 13, 5
					transmit
					call write_value_and_transmit("RPPH", 1, 8)
				End If
				If disenrollment_reason <> "" Then
                    EMReadScreen beg_of_curr_span, 8, 13, 5
                    If beg_of_curr_span = enrollment_date Then
                        EMWriteScreen "...", 13, 5
                    Else
                        EMWriteScreen xcl_end_date, 13, 14
                        EMWriteScreen disenrollment_reason, 13, 75
                    End If
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
				' MsgBox "STOP HERE"
    			EMWaitReady 0, 0

    			EMReadScreen false_end, 8, 14, 14
    			If false_end = "99/99/99" Then
    				EMReadScreen double_check, 2, 14, 5
    				If double_check = "  " Then EMWriteScreen "...", 14, 5
    			End If
    			'msgbox "RPPH updated"

    			'REFM screen
    			' EMWriteScreen "refm", 1, 8
    			transmit
				EMReadScreen current_panel, 4, 1, 52
    			EMReadScreen RPPH_error_check, 78, 24, 2

				RPPH_error_check = trim(RPPH_error_check)
				enrollment_successful = TRUE

				' MsgBox "Panel - " & current_panel & vbCr & "Error - " & RPPH_error_check
				If current_panel <> "REFM" Then
					enrollment_successful = FALSE
				' ElseIf RPPH_error_check <> "" Then

				Else
					'updating REFM'
					call enter_detail_on_refm

					'nav to RPPH
					EMWriteScreen "rpph", 1, 8
					transmit

					EMReadScreen panel_enrollment_date, 8, 13, 5
					'enter managed care plan code
					EMReadScreen panel_health_plan_code, 10, 13, 23
					'enter contract code
					EMReadScreen panel_contract_code_part_one, 2, 13, 34
					EMReadScreen panel_contract_code_part_two, 2, 13, 37
					'enter change reason
					EMReadScreen panel_change_reason, 2, 13, 71

					If panel_enrollment_date <> enrollment_date Then enrollment_successful = FALSE
					If panel_health_plan_code <> health_plan_code Then enrollment_successful = FALSE
					If panel_contract_code_part_one <> contract_code_part_one Then enrollment_successful = FALSE
					If panel_contract_code_part_two <> contract_code_part_two Then enrollment_successful = FALSE
					If panel_change_reason <> change_reason Then enrollment_successful = FALSE
					' MsgBox enrollment_successful
					If enrollment_successful = TRUE Then
						EMWriteScreen "refm", 1, 8
		                transmit
					End If
				End If

				If enrollment_successful = FALSE Then
					testing_run = TRUE
					If new_enrol_date = "" Then new_enrol_date = enrollment_date
					new_enrol_date = new_enrol_date & ""
					script_run_lowdown = script_run_lowdown & "Enrollment was not successful for " & MMIS_clients_array(client_name, member) & vbCr
					' MsgBox "~" & new_enrol_date & "~"
					If MMIS_clients_array(new_plan_two, member) = "" Then MMIS_clients_array(new_plan_two, member) = MMIS_clients_array(new_plan, member)
					If MMIS_clients_array(contr_code_two, member) = "" Then MMIS_clients_array(contr_code_two, member) = MMIS_clients_array(contr_code, member)
					If MMIS_clients_array(change_rsn_two, member) = "" Then MMIS_clients_array(change_rsn_two, member) = MMIS_clients_array(change_rsn, member)
					If MMIS_clients_array(disenrol_rsn_two, member) = "" Then MMIS_clients_array(disenrol_rsn_two, member) = MMIS_clients_array(disenrol_rsn, member)

					Do
						err_msg = ""

						BeginDialog Dialog1, 0, 0, 456, 290, "Update Enrollment Options due to MMIS Failure or Error"
						  DropListBox 80, 185, 125, 45, "Select One..."+chr(9)+"Blue Plus"+chr(9)+"Health Partners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Ucare"+chr(9)+"United Healthcare", MMIS_clients_array(new_plan_two, member)
						  DropListBox 250, 185, 65, 45, "Select One..."+chr(9)+"MA 12"+chr(9)+"NM 12"+chr(9)+"MA 30"+chr(9)+"MA 35"+chr(9)+"IM 12", MMIS_clients_array (contr_code_two, member)
						  EditBox 390, 185, 50, 15, new_enrol_date
						  DropListBox 80, 205, 130, 45, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", MMIS_clients_array(change_rsn_two, member)
						  DropListBox 100, 225, 130, 45, "Select one..."+chr(9)+"Eligibility ended"+chr(9)+"Exclusion"+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Jail - Incarceration"+chr(9)+"Move"+chr(9)+"Loss of disability"+chr(9)+"Ninety Day change option"+chr(9)+"Open Enrollment"+chr(9)+"PMI merge"+chr(9)+"Voluntary", MMIS_clients_array(disenrol_rsn_two, member)
						  ButtonGroup ButtonPressed
						    PushButton 10, 270, 135, 15, "Skip Enrollment for " & MMIS_clients_array(client_name, member), skip_enrollment_btn
						    PushButton 160, 270, 150, 15, "Enrollment Already Updated Manually", read_manual_enrollment_btn
						    PushButton 325, 270, 125, 15, " Retry Enrollment with Script", retry_enrollment_btn
						  Text 10, 10, 345, 10, "The selections made for this person caused a message or error in MMIS and were not able to be updated."
						  GroupBox 10, 25, 440, 105, "Initial Selections for " & MMIS_clients_array(client_name, member)
						  Text 20, 45, 190, 10, "Plan to Enroll:" & MMIS_clients_array(new_plan, member)
						  Text 215, 45, 75, 10, "ID / Desc:" & MMIS_clients_array (contr_code, member)
						  Text 330, 45, 110, 10, "Enrollment Date:" & enrollment_date
						  Text 20, 60, 210, 10, "Change Reason:" & MMIS_clients_array(change_rsn, member)
						  Text 20, 75, 220, 10, "Disenrollment Reason:" & MMIS_clients_array(disenrol_rsn, member)
						  Text 20, 90, 55, 10, "Pregnant: " & pregnant_yn
						  Text 115, 90, 60, 10, "Interpreter: " & interpreter_yn
						  Text 190, 90, 90, 10, "Interpreter Code:" & MMIS_clients_array(interp_code, member)
						  Text 20, 115, 215, 10, "Current Enrollment:" & MMIS_clients_array(current_plan, member)
						  GroupBox 10, 135, 440, 30, "MMIS Error Messages"
						  Text 20, 150, 415, 10, RPPH_error_check
						  GroupBox 10, 170, 440, 95, "Change Selections for " & MMIS_clients_array(client_name, member)
						  Text 20, 190, 60, 10, "Enrollment Plan:"
						  Text 215, 190, 35, 10, "ID / Desc:"
						  Text 330, 190, 60, 10, "Enrollment Date:"
						  Text 20, 210, 60, 10, "Change Reason: "
						  Text 20, 230, 80, 10, "Disenrollment Reason:"
						EndDialog

						dialog Dialog1
						cancel_confirmation

						If ButtonPressed = retry_enrollment_btn Then

							If MMIS_clients_array(new_plan_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "*  Select which plan you need to enroll " & MMIS_clients_array(client_name, person) & "into."
							If MMIS_clients_array (contr_code_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the product ID and description for enrollment."
							If MMIS_clients_array(change_rsn_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select a reason to enroll  " & MMIS_clients_array(client_name, person) & " into a new plan."
							If left(MMIS_clients_array(current_plan, person), 3) <> "XCL" AND trim(MMIS_clients_array(current_plan, person)) <> "" Then
								If MMIS_clients_array(disenrol_rsn_two, member) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since " & MMIS_clients_array(client_name, person) & " is currently on a health plan, please select a disenrollment reason for the " & MMIS_clients_array(current_plan, person) & " plan."
							End If
							If err_msg <> "" THen MsgBox "Please resovle to continue:" & vbNewLine & err_msg

						End If
					Loop until err_msg = ""

					new_enrol_date = DateAdd("d", 0, new_enrol_date)
					If ButtonPressed = retry_enrollment_btn Then
						PF10
						MMIS_clients_array(new_plan, member) = MMIS_clients_array(new_plan_two, member)
						MMIS_clients_array(contr_code, member) = MMIS_clients_array(contr_code_two, member)
						MMIS_clients_array(change_rsn, member) = MMIS_clients_array(change_rsn_two, member)
						MMIS_clients_array(disenrol_rsn, member) = MMIS_clients_array(disenrol_rsn_two, member)

						enrollment_date = new_enrol_date
						script_run_lowdown = script_run_lowdown & "Button Pressed - RETRY ENROLLMENT" & vbCr
					End If
					If ButtonPressed = read_manual_enrollment_btn Then
						EMReadScreen current_panel, 4, 1, 52

						If current_panel = "RPPH" Then
							Call read_detail_on_RPPH(MMIS_clients_array(manual_enrollment_date, member), panel_health_plan_code, panel_contract_code_part_one, panel_contract_code_part_two, MMIS_clients_array(manual_contr_code, member), panel_change_reason, panel_disenrollment_reason, MMIS_clients_array(manual_new_plan, member), MMIS_clients_array(manual_change_rsn, member), MMIS_clients_array(manual_disenrol_rns, member))

							EMWriteScreen "refm", 1, 8
			                transmit

							call enter_detail_on_refm
						ElseIf current_panel = "REFM" Then
							call enter_detail_on_refm

							EMWriteScreen "rpph", 1, 8
							transmit

							Call read_detail_on_RPPH(MMIS_clients_array(manual_enrollment_date, member), panel_health_plan_code, panel_contract_code_part_one, panel_contract_code_part_two, MMIS_clients_array(manual_contr_code, member), panel_change_reason, panel_disenrollment_reason, MMIS_clients_array(manual_new_plan, member), MMIS_clients_array(manual_change_rsn, member), MMIS_clients_array(manual_disenrol_rns, member))
						End If
						enrollment_successful = TRUE
						MMIS_clients_array(manual_enrollment, member) = TRUE
						PF9
						Do
							PF3
							EMReadScreen where_are_we, 4, 1, 52
						Loop until where_are_we = "RKEY"
						script_run_lowdown = script_run_lowdown & "Button Pressed - MANUAL ENROLLMENT" & vbCr
					End If
					If ButtonPressed = skip_enrollment_btn Then
						process_manually_message = process_manually_message & "Enrollment was cancelled by you when you pressed the button 'Skip Enrollment for " & MMIS_clients_array(client_name, member) & "' button on the 'Update Enrollment Options due to MMIS Failure or Error' dialog." & vbNewLine & vbNewLine & ""
						enrollment_successful = TRUE
						PF10
						script_run_lowdown = script_run_lowdown & "Button Pressed - SKIP ENROLLMENT" & vbCr
					End If

					first_attempt = FALSE
				End If
            ELSE
                'REFM screen
                EMWriteScreen "refm", 1, 8
                transmit
            End If

			'blanking out varibles if the other option is selected
			If change_reason = "Select one..." then change_reason = ""
			If disenrollment_reason = "Select one..." then disenrollment_reason = ""
			'making sure script got to right panel
			' EMReadScreen REFM_check, 4, 1, 52
			' If REFM_check <> "REFM" then process_manually_message = process_manually_message & "The script was unable to navigate to REFM for PMI " & MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & "needs to be processed manually." & vbNewLine & vbNewLine
		Loop until enrollment_successful = TRUE

		If MMIS_clients_array(manual_enrollment, member) <> TRUE Then
			PF9
			If MMIS_clients_array(current_plan, member) = "XCL - Adoption Assistance" Then PF9

			'error handling to ensure that enrollment date and exclusion dates don't conflict
			EMReadScreen REFM_error_check, 19, 24, 2 'checks for an inhibiting edit
			IF REFM_error_check <> "                   " then
	            IF REFM_error_check <> "INVALID KEY ENTERED" AND REFM_error_check <> "INVALID KEY PRESSED" then
	                EMReadScreen full_error_msg, 79, 24, 2
	                full_error_msg = trim(full_error_msg)
				    process_manually_message = process_manually_message & "You have entered information that is causing a warning error, or an inhibiting error for PMI "& MMIS_clients_array(client_pmi, member) & ". The enrollment for " & MMIS_clients_array(client_name, member) & ". Refer to the MMIS USER MANUAL to resolve if necessary. Full error message: " & full_error_msg & vbNewLine & vbNewLine
			    END IF
	        END IF
		End If
		' msgbox "all updated - see casenote code"
		'Save and case note
		EMReadScreen where_are_we, 4, 1, 52
		Do While where_are_we <> "RKEY"
			pf3
			EMReadScreen where_are_we, 4, 1, 52
		Loop

		EMWriteScreen "i", 2, 19
        EMWriteScreen "        ", 9, 19
        EMWriteScreen MMIS_clients_array(client_pmi, member), 4, 19
		transmit
		check_enrollment = TRUE
		IF ButtonPressed = read_manual_enrollment_btn Then check_enrollment = FALSE
		IF ButtonPressed = skip_enrollment_btn Then check_enrollment = FALSE
		EMReadScreen rsum_enrollment, 8, 16, 20
		EMReadScreen rsum_plan, 10, 16, 52

		' MsgBox "RSUM date and plan: " & rsum_enrollment & " - " & rsum_plan & vbNewLine & "Coded date and plan: " & Enrollment_date & " - " & health_plan_code
		IF rsum_enrollment = Enrollment_date AND rsum_plan = health_plan_code Then
			MMIS_clients_array(enrol_sucs, member) = TRUE
			script_run_lowdown = script_run_lowdown & "Enrollment appears successful for " & MMIS_clients_array(client_name, member) & vbCr
		''			pf4
		''			pf11
		''			EMSendkey "***HMO Note*** " &  MMIS_clients_array(client_name, member) & " enrolled into " &  MMIS_clients_array(new_plan, member) & " " & Enrollment_date & " " & worker_signature
		''			pf3
		Else
			failed_enrollment_message = failed_enrollment_message & vbNewLine & vbNewLine & process_manually_message
			MMIS_clients_array(enrol_sucs, member) = FALSE
			script_run_lowdown = script_run_lowdown & "Enrollment appears to have failed for " & MMIS_clients_array(client_name, member) & vbCr & "Process manually message:" & vbCr & process_manually_message & "-----------------" & vbCr
		End If
		' MsgBox process_manually_message
		pf3
		IF REFM_error_check = "WARNING: MA12,01/16" Then
			PF3
		END IF
		process_manually_message = ""
	Next
End If

If enrollment_source = "Morning Letters" Then
	name_of_script = "ACTIONS - MHC ENROLLMENT - MORN LTRS.vbs"
	If open_enrollment_case = TRUE Then name_of_script = "ACTIONS - MHC AHPS ENROLLMENT - MORN LTRS.vbs"
Else
	name_of_script = "ACTIONS - MHC ENROLLMENT - " & UCASE(left(enrollment_source, 5)) & ".vbs"
	If open_enrollment_case = TRUE Then name_of_script = "ACTIONS - MHC AHPS ENROLLMENT - " & UCASE(left(enrollment_source, 5)) & ".vbs"
End If
If caller_rela = "" Then caller_rela = "Client"

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

create_case_note = FALSE
For member = 0 to Ubound(MMIS_clients_array, 2)
	If MMIS_clients_array(enrol_sucs, member) = TRUE Then create_case_note = TRUE
Next

' CALL write_variable_in_MMIS_NOTE ("***Hennepin MHC note*** Household enrollment updated for " & Enrollment_date & " per enrollment form")
If create_case_note = TRUE Then
	If open_enrollment_case = TRUE Then
		CALL write_variable_in_MMIS_NOTE ("AHPS request processed for 2022 selection")
		If enrollment_source = "Morning Letters" Then
		ElseIf enrollment_source = "Phone" Then
			CALL write_variable_in_MMIS_NOTE ("Enrollment requested by " & caller_rela & " via " & enrollment_source)
		ElseIf enrollment_source = "Paper Enrollment Form" Then
			CALL write_variable_in_MMIS_NOTE ("Enrollment requested via " & enrollment_source)
		End If
	Else
		If enrollment_source = "Morning Letters" Then
		    CALL write_variable_in_MMIS_NOTE ("Re-enrollment processed effective: " & enrollment_date)
		    CALL write_variable_in_MMIS_NOTE ("Following clients had PMAP under duplicate PMI(s) in the last 12 months:")
		ElseIf enrollment_source = "Phone" Then
		    CALL write_variable_in_MMIS_NOTE ("Enrollment effective: " & enrollment_date & " requested by " & caller_rela & " via " & enrollment_source)
		ElseIf enrollment_source = "Paper Enrollment Form" Then
		    CALL write_variable_in_MMIS_NOTE ("Enrollment effective: " & enrollment_date & " requested via " & enrollment_source)
		End If
	End If
	If enrollment_source = "Phone" Then CALL write_variable_in_MMIS_NOTE("Call completed " & now & " with " & caller_name & " from the number: " & phone_number_of_caller)
	If used_interpreter_checkbox = checked then CALL write_variable_in_MMIS_NOTE("Interpreter used for phone call.")
	If trim(form_received_date) <> "" Then CALL write_variable_in_MMIS_NOTE("Enrollment requested via Form received on " & form_received_date & ".")
	For member = 0 to Ubound(MMIS_clients_array, 2)
		If MMIS_clients_array(enrol_sucs, member) = TRUE Then
	        If enrollment_source = "Morning Letters" Then
	            If MMIS_clients_array(manual_enrollment, member) = FALSE Then CALL write_variable_in_MMIS_NOTE ("- Re-enrolled " & MMIS_clients_array(client_name, member) & " in " & MMIS_clients_array(new_plan, member))
				If MMIS_clients_array(manual_enrollment, member) = TRUE  Then CALL write_variable_in_MMIS_NOTE ("- Re-enrolled " & MMIS_clients_array(client_name, member) & " in " & MMIS_clients_array(manual_new_plan, member) & " effective: " & MMIS_clients_array(manual_enrollment_date, member))
	        Else
				If MMIS_clients_array(manual_enrollment, member) = FALSE Then CALL write_variable_in_MMIS_NOTE ("- " & MMIS_clients_array(client_name, member) & " enrolled into " & MMIS_clients_array(new_plan, member))
			    If MMIS_clients_array(manual_enrollment, member) = TRUE  Then CALL write_variable_in_MMIS_NOTE ("- " & MMIS_clients_array(client_name, member) & " enrolled into " & MMIS_clients_array(manual_new_plan, member) & " effective: " & MMIS_clients_array(manual_enrollment_date, member))
	        End If
		End If
	Next
	CALL write_bullet_and_variable_in_MMIS_NOTE ("Notes", other_notes)

	CALL write_variable_in_MMIS_NOTE ("Processed by " & worker_signature)
	CALL write_variable_in_MMIS_NOTE ("*************************************************************************")
	pf3
	pf3
	IF REFM_error_check = "WARNING: MA12,01/16" Then
		PF3
	END IF
End If

If create_case_note = FALSE Then
	failed_enrollment_message = "Script run is complete but NO NOTE has been entered in MMIS as none of the members were able to be enrolled." & vbNewLine & vbNewLine & "Members that could not be enrolled:" & vbNewLine & vbNewLine & "*****" & vbNewLine & failed_enrollment_message
Else
	If trim(failed_enrollment_message) = "" Then
		failed_enrollment_message = "The script is complete. Enrollment has been updated and case noted." & vbNewLine & vbNewLine & "It appears all clients were able to be enrolled as requested."
	Else
		failed_enrollment_message = "The script is complete. Enrollment has been updated and case noted." & vbNewLine & vbNewLine & "Some clients enrollments could not be processed by the script for some reason, they are listed below:" & vbNewLine & vbNewLine & "*****" & vbNewLine & failed_enrollment_message
	End If
End If

script_end_procedure_with_error_report (failed_enrollment_message)


' BeginDialog Dialog1, 0, 0, 376, 265, "Dialog"
'   GroupBox 10, 5, 360, 90, "Requested Enrollment"
'   Text 20, 20, 340, 10, "You have selected the following enrollment for CLIENT NAME GOES HERE"
'   Text 30, 35, 190, 10, "Health Plan: SELECTED HEALTH PLAN"
'   Text 30, 45, 185, 10, "Contract code: MA 12"
'   Text 30, 55, 185, 10, "Change Reason:"
'   Text 30, 65, 185, 10, "Disenroll Reason:"
'   Text 30, 75, 185, 10, "Enrollment Effective Date 01/2021"
'   GroupBox 10, 100, 360, 60, "Enrollment Failed"
'   Text 20, 115, 310, 10, "The enrollment failed with the above selections. Message(s) from MMIS read:"
'   Text 30, 130, 325, 10, "MMIS ERROR CODE AND MESSAGE GO HERE"
'   Text 30, 145, 325, 10, "MMIS ERROR CODE AND MESSAGE GO HERE"
'   GroupBox 10, 165, 360, 75, "New Selection Information for CLIENT NAME"
'   Text 35, 185, 50, 10, "Health Plan:"
'   DropListBox 80, 180, 145, 45, "", List1
'   Text 235, 185, 50, 10, "Contract Code:"
'   DropListBox 290, 180, 75, 45, "", List2
'   Text 25, 205, 50, 10, "Cange Reason:"
'   DropListBox 80, 200, 145, 45, "", List3
'   Text 15, 225, 65, 15, "Disenroll Reason:"
'   DropListBox 80, 220, 145, 45, "", List4
'   Text 235, 205, 70, 10, "Enrollment MM/YY"
'   EditBox 310, 200, 25, 15, Edit1
'   EditBox 340, 200, 25, 15, Edit2
'   ButtonGroup ButtonPressed
'     PushButton 10, 245, 100, 15, "Cancel this Enrollment", Button3
'     PushButton 220, 245, 150, 15, "Try Enrollment Again with New Selections", Button5
' EndDialog
