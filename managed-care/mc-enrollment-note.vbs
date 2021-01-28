'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MHC NOTE ENROLLMENT.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 60			 'manual run time in seconds
STATS_denomination = "M"		 'M is for Member
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
call changelog_update("04/24/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'call check_for_MMIS(True) 'Sending MMIS back to the beginning screen and checking for a password prompt
Call get_to_RKEY

'grabs the PMI number if one is listed on RKEY
EMReadscreen MMIS_case_number, 8, 9, 19
MMIS_case_number= trim(MMIS_case_number)

enrollment_month = CM_plus_1_mo
enrollment_year = CM_plus_1_yr

this_month = monthname(month(date))
Select Case this_month
    Case "January"
        cut_off_date = #01/22/2020#
    Case "February"
        cut_off_date = #2/19/2020#
    Case "March"
        cut_off_date = #3/20/2020#
    Case "April"
        cut_off_date = #4/21/2020#
    Case "May"
        cut_off_date = #5/19/2020#
    Case "June"
        cut_off_date = #6/19/2020#
    Case "July"
        cut_off_date = #7/22/2020#
    Case "August"
        cut_off_date = #8/20/2020#
    Case "September"
        cut_off_date = #9/21/2020#
    Case "October"
        cut_off_date = #10/21/2020#
    Case "November"
        cut_off_date = #11/17/2020#
    Case "December"
        cut_off_date = #12/21/2020#
End Select
'MsgBox cut_off_date
If cut_off_date <> "" Then
    If DateDiff("d", date, cut_off_date) < 0 Then
        'MsgBox DateDiff("d", date, cut_off_date)
        enrollment_month = CM_plus_2_mo
        enrollment_year = CM_plus_2_yr
    End If
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 206, 110, "Enrollment Information"
  EditBox 90, 25, 60, 15, MMIS_case_number
  EditBox 90, 45, 25, 15, enrollment_month
  EditBox 120, 45, 25, 15, enrollment_year
  DropListBox 110, 70, 90, 45, "Select One..."+chr(9)+"Phone"+chr(9)+"Paper Enrollment Form", enrollment_source
  ButtonGroup ButtonPressed
    OkButton 95, 90, 50, 15
    CancelButton 150, 90, 50, 15
  GroupBox 5, 10, 150, 55, "Leading zeros not needed"
  Text 10, 30, 50, 10, "Case Number:"
  Text 10, 50, 80, 10, "Enrollment Month/Year:"
  Text 5, 70, 100, 10, "Enrollment was requested via"
EndDialog

'do the dialog here
Do
    err_msg = ""

	Dialog Dialog1
	cancel_confirmation

	If trim(MMIS_case_number) = "" then err_msg = err_msg & vbNewLine & "* Enter the case number."
    If enrollment_month = "" OR enrollment_year = "" Then err_msg = err_msg & vbNewLine & "* Enter the month and year enrollment is effective."
    If enrollment_source = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate where the request for the enrollment came from (phone call or enrollment form)."

    If err_msg <> "" Then MsgBOx "Please resolve to continue: " & vbNewLine & err_msg
Loop until err_msg = ""

MMIS_case_number = trim(MMIS_case_number)

'checking for an active MMIS session
Call check_for_MMIS(True)
Call get_to_RKEY

'formatting variables----------------------------------------------------------------------------------------------------
If len(enrollment_month) = 1 THEN enrollment_month = "0" & enrollment_month
IF len(enrollment_year) <> 2 THEN enrollment_year = right(enrollment_year, 2)

MNSURE_Case = False
If len(MMIS_case_number) = 8 AND left(MMIS_case_number, 1) <> 0 THEN MNSURE_Case = TRUE
MMIS_case_number = right("00000000" & MMIS_case_number, 8)

enrollment_date = enrollment_month & "/01/" & enrollment_year

EMWriteScreen "i", 2, 19
EMWriteScreen MMIS_case_number, 9, 19
transmit
transmit
transmit
EMReadscreen RCIN_check, 4, 1, 49
If RCIN_check <> "RCIN" then script_end_procedure("The listed Case number was not found. Check your Case number and try again.")

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

Interim_array = split(client_array, "|")

PF7
PF7

DIM all_clients_array()
ReDim all_clients_array(total_clients, 1)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	all_clients_array(x, 0) = Interim_array(x)
    all_clients_array(x, 1) = 1
NEXT

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
const contract_code     = 6
const first_name_ini    = 7
const contr_code   = 8
const preg_yes 	   = 9
const interp_code  = 10
const enrol_sucs   = 11
const case_note_checkbox = 12

Dim MMIS_clients_array
ReDim MMIS_clients_array (12, 0)

item = 0

For each member in HH_member_array

    EMReadScreen RCIN_check, 4, 1, 49
    If RCIN_check = "RCIN" Then PF6
    Call get_to_RKEY

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
    MMIS_clients_array (first_name_ini, item) = first_name & " " & left(last_name, 1) & "."

    EMWriteScreen "RPPH", 1, 8
	transmit

    row = 1
	col = 1
	EMSearch "99/99/99", row, col
    IF row > 10 Then

        EMReadScreen begin_enrollment, 8, row, 5

        If begin_enrollment = enrollment_date Then
            MMIS_clients_array(case_note_checkbox, item) = checked

            EMReadScreen hp_code, 10, row, 23

            If hp_code = "A585713900" then MMIS_clients_array(current_plan, item) = "HealthPartners"
            If hp_code = "A565813600" then MMIS_clients_array(current_plan, item) = "Ucare"
            If hp_code = "A405713900" then MMIS_clients_array(current_plan, item) = "Medica"
            If hp_code = "A065813800" then MMIS_clients_array(current_plan, item) = "BluePlus"
            If hp_code = "A836618200" then MMIS_clients_array(current_plan, item) = "Hennepin Health PMAP"
            If hp_code = "A965713400" then MMIS_clients_array(current_plan, item) = "Hennepin Health SNBC"

            EMReadScreen plan_id, 5, row, 34

            MMIS_clients_array(contract_code, item) = plan_id

            EMReadScreen chg_rsn_code, 2, row, 71

            If chg_rsn_code = "AP" Then MMIS_clients_array(change_rsn, item) = "Appeal"
            If chg_rsn_code = "FY" Then MMIS_clients_array(change_rsn, item) = "First year change option"
            If chg_rsn_code = "HP" Then MMIS_clients_array(change_rsn, item) = "Health plan contract end"
            If chg_rsn_code = "IN" Then MMIS_clients_array(change_rsn, item) = "Initial enrollment"
            If chg_rsn_code = "MV" Then MMIS_clients_array(change_rsn, item) = "Move"
            If chg_rsn_code = "NT" Then MMIS_clients_array(change_rsn, item) = "Ninety Day change option"
            If chg_rsn_code = "OE" Then MMIS_clients_array(change_rsn, item) = "Open enrollment"
            If chg_rsn_code = "OT" Then MMIS_clients_array(change_rsn, item) = "Other"
            If chg_rsn_code = "PM" Then MMIS_clients_array(change_rsn, item) = "PMI merge"
            If chg_rsn_code = "RE" Then MMIS_clients_array(change_rsn, item) = "Reenrollment"
            If chg_rsn_code = "RS" Then MMIS_clients_array(change_rsn, item) = "Reinstatement"
            If chg_rsn_code = "SE" Then MMIS_clients_array(change_rsn, item) = "Service Ending"
            If chg_rsn_code = "VL" Then MMIS_clients_array(change_rsn, item) = "Voluntary"

        End If

    End If

    item = item + 1
    PF3

Next

x = 0
max = Ubound(MMIS_clients_array, 2)
dlg_len = 80
y_pos = 5
If enrollment_source = "Phone" Then
    dlg_len = dlg_len + 40
    y_pos = 45
End If

name_list = ""
For person = 0 to Ubound(MMIS_clients_array, 2)
    name_list = name_list & +chr(9)+MMIS_clients_array(first_name_ini, person)
Next

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 476, (max * 20) + dlg_len, "Enrollment Information"

  If enrollment_source = "Phone" Then
      GroupBox 5, 0, 460, 35, "Phone Call Information"
      Text 10, 20, 40, 10, "Caller name"
      ComboBox 55, 15, 120, 45, " " & name_list, caller_name
      Text 180, 20, 40, 10, ", who is the"
      ComboBox 225, 15, 80, 45, "Client"+chr(9)+"AREP", caller_rela
      CheckBox 390, 15, 65, 10, "Used Interpreter", used_interpreter_checkbox
  End If

  Text 5, y_pos, 30, 10, "Include?"
  Text 40, y_pos, 25, 10, "Name"
  Text 135, y_pos, 15, 10, "PMI"
  Text 180, y_pos, 60, 10, "Plan Enrolled Into"
  Text 365, y_pos, 55, 10, "Change reason:"
  y_pos = y_pos + 20

  For person = 0 to Ubound(MMIS_clients_array, 2)
    CheckBox 5, (x * 20) + y_pos, 25, 10, "Yes", MMIS_clients_array(case_note_checkbox, person)
  	Text 40, (x * 20) + y_pos, 95, 10, MMIS_clients_array(client_name, person)
  	Text 135, (x * 20) + y_pos, 35, 10, MMIS_clients_array(client_pmi, person)
    DropListBox 180, (x * 20) + y_pos - 5, 105, 15, " "+chr(9)+"BluePlus"+chr(9)+"HealthPartners"+chr(9)+"Hennepin Health PMAP"+chr(9)+"Medica"+chr(9)+"Hennepin Health SNBC"+chr(9)+"Ucare", MMIS_clients_array(current_plan, person)
  	DropListBox 295, (x * 20) + y_pos - 5, 40, 15, "MA 12"+chr(9)+"NM 12"+chr(9)+"MA 30"+chr(9)+"MA 35"+chr(9)+"MA 37", MMIS_clients_array(contr_code, person)
	DropListBox 365, (x * 20) + y_pos - 5, 105, 15, "Select one..."+chr(9)+"First year change option"+chr(9)+"Health plan contract end"+chr(9)+"Initial enrollment"+chr(9)+"Move"+chr(9)+"Ninety Day change option"+chr(9)+"Open enrollment"+chr(9)+"PMI merge"+chr(9)+"Reenrollment", MMIS_clients_array(change_rsn, person)
	x = x + 1
  Next
  y_pos = y_pos + 20

  Text 5, (max * 20) + y_pos, 45, 10, "Other Notes:"
  EditBox 55, (max * 20) + y_pos - 5, 415, 15, other_notes
  y_pos = y_pos + 20
  Text 5, (max * 20) + y_pos, 60, 10, "Worker Signature"
  EditBox 70, (max * 20) + y_pos - 5, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 365, (max * 20) + y_pos - 5, 50, 15
    CancelButton 420, (max * 20) + y_pos - 5, 50, 15
EndDialog

Call get_to_RKEY
EMWriteScreen "i", 2, 19
EMWriteScreen MMIS_case_number, 9, 19
transmit
transmit
transmit

Do
    err_msg = ""

	Dialog Dialog1
	cancel_confirmation

    member_selected = FALSE
    For person = 0 to Ubound(MMIS_clients_array, 2)
        If MMIS_clients_array(case_note_checkbox, person) = checked Then member_selected = TRUE
    Next

    If enrollment_source = "Phone" Then

        If trim(caller_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the caller."
        If trim(caller_rela) = "" Then err_msg = err_msg & vbNewLine & "* Select who is calling (typically Client or AREP)."

    End If

    If member_selected = FALSE Then err_msg = err_msg & vbNewLine & "* You must select at least one person that had an enrollment processed."
    If worker_signature = "" THen err_msg = err_msg & vbNewLine & "* Enter your name for the case note signature."

    If err_msg <> "" THen MsgBox "Please resovle to continue:" & vbNewLine & err_msg

Loop Until err_msg = ""

name_of_script = "NOTES - MHC NOTE ENROLLMENT - " & left(enrollment_source, 5) & ".vbs"
If caller_rela = "" Then caller_rela = "Client"

Call check_for_MMIS(False)
Call get_to_RKEY

For person = 0 to Ubound(MMIS_clients_array, 2)
    If MMIS_clients_array(case_note_checkbox, person) = checked Then
        case_note_pmi = MMIS_clients_array(client_pmi, person)
        Exit For
    End If
Next

'Case Noting - goes into RSUM for the first client to do the case note
EMWriteScreen "c", 2, 19
EMWriteScreen "        ", 9, 19
EMWriteScreen case_note_pmi, 4, 19
transmit
pf4
pf11		'Starts a new case note'

CALL write_variable_in_MMIS_NOTE ("Enrollment effective: " & enrollment_date & " requested by " & caller_rela & " via " & enrollment_source)
row = 7
If enrollment_source = "Phone" Then CALL write_variable_in_MMIS_NOTE("Call completed " & now & " with " & caller_name)
If used_interpreter_checkbox = checked then CALL write_variable_in_MMIS_NOTE("Interpreter used for phone call.")
For member = 0 to Ubound(MMIS_clients_array, 2)
	If MMIS_clients_array(case_note_checkbox, member) = checked Then
		CALL write_variable_in_MMIS_NOTE (MMIS_clients_array(client_name, member) & " enrolled into " & MMIS_clients_array(current_plan, member))
        STATS_counter = STATS_counter + 1
		row = row + 1
	End If
Next
CALL write_bullet_and_variable_in_MMIS_NOTE ("Notes", other_notes)
row = row + 1
CALL write_variable_in_MMIS_NOTE ("Processed by " & worker_signature)
CALL write_variable_in_MMIS_NOTE ("*************************************************************************")

'MsgBox "Review"
PF3 'Leaving edit mode
PF4 'Going back to see case note

' pf3
' pf3
' IF REFM_error_check = "WARNING: MA12,01/16" Then
' 	PF3
' END IF

MAXIS_case_number = MMIS_case_number

' 'REMOVING ESSO CODE AS WE ARE UPGRADING BZ AND A BUCH OF STUFF _ WILL REVIEW LATER
' 'End of script code for restarting ESSO
' IF using_ESSO = TRUE THEN
'   'MsgBox "End of script reached. Because ESSO was previously found on your computer, attempting to start ESSO in the background..."
'   SET ObjShell = CreateObject("Wscript.Shell")
'   ObjShell.Run """C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Oracle\ESSO-LM\ESSO-LM.lnk"""
'   vgo_msg = "ESSO started, the ESSO icon should be added back to the system tray."
' ELSE
'   vgo_msg = "End of script reached. Because ESSO was not previously found on your computer, there is no need to try to start ESSO."
' END IF
'
' end_msg = "Success! NOTE entered in to MMIS of enrollment processed." &vbNewLine & vbNewLine & vgo_msg
' script_end_procedure(end_msg)

script_end_procedure("Success! NOTE entered in to MMIS of enrollment processed.")
