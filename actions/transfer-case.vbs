'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TRANSFER CASE.vbs"
start_time = timer
STATS_counter = 1                  	'sets the stats counter at one
STATS_manualtime = 0              'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("06/05/2026", "Major change: New internal transfer process will select correct transfer caseload based on information entered in dialog. See script instructions for full details.", "Dave Courtright, Hennepin County.")
Call changelog_update("03/12/2024", "Updated resource links.", "Megan Geissler, Hennepin County.")
Call changelog_update("01/23/2024", "Added reminder to inner-county option to transfer cases in ECF Next prior to transferring the case in MAXIS.", "Ilse Ferris, Hennepin County.")
call changelog_update("12/15/2023", "Updated resource links.", "Megan Geissler, Hennepin County.")
call changelog_update("11/07/2023", "Fixed bug when case is a MA Excluded Time case.", "Ilse Ferris, Hennepin County.")
call changelog_update("05/24/2023", "Updated date validation and error messages for dialog to ensure consistent entry.", "Mark Riegel, Hennepin County.")
call changelog_update("10/26/2022", "Updated several functionalities to support enhanced experience for both inner and inter county transfer.", "Ilse Ferris, Hennepin County.")
call changelog_update("04/01/2022", "There is no longer a case note for in-county transfers based on guidance provided by CASE NOTE III: CLAIMS/SYSTEMS/TRANSFERS TE02.08.095.", "MiKayla Handley, Hennepin County.")
call changelog_update("03/28/2022", "Multiple updates made ensuring that the transfer is complete and removing the case from in-county transfers.", "MiKayla Handley, Hennepin County.")
CALL changelog_update("05/21/2021", "Updated browser to default when opening SIR from Internet Explorer to Edge.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/09/2020", "No issues with SPEC/MEMO for out-of-county cases. SIR Announcement from 11/05/20 stated an issue was identified. Hennepin County's script project is seperate from DHS's script project. We are not experiencing the reported issue. Thank you!", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/20/2020", "Updated link for out-of-county use form.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/16/2020", "Added closing message box with information if a case is identified as being transferred to the MNPrairie Bank of counties.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/20/2020", "Rewrite.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/19/2019", "Added an error reporting option at the end of the script run.", "Casey Love, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'--------------------------------------------------------------------------------The script
EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
Call Check_for_MAXIS(false)                         'Ensuring we are in MAXIS


Dim transfer_to_worker
alpha_split_one_a_h = "ABCDEFGH"
alpha_split_two_i_z = "IJKLMNOPQRSTUVWXYZ"
alpha_split_one_a_l = "ABCDEFGHIJKL"
alpha_split_two_m_z = "MNOPQRSTUVWXYZ"

'Load the caseload-directory file from misc
Call run_from_GitHub(script_repository & "\misc\caseload-directory.vbs")

'Functions from Application received
function select_random_index(ubound_of_array, index_selection)
	If ubound_of_array = 0 Then
		index_selection = 0
	Else
		options = ubound_of_array + 1
		Randomize       'Before calling Rnd, use the Randomize statement without an argument to initialize the random-number generator.
		rnd_nbr = rnd
		size_up = rnd_nbr * options
		index_selection = int(size_up)
	End If
end function

function get_caseload_array_by_type(caseload_type, available_caseload_array)
	all = caseload_info.items
	things = caseload_info.keys

	Dim temp_array()
	ReDim temp_array(0)
	counter = 0
    random_team_needed = False
	for i = 0 to UBound(all)
        If right(caseload_type, 1) = "?" Then random_team_needed = True 'failsafe to ensure that the random team is selected when something slips through the cracks
        If random_team_needed = TRUE Then  'This will be used to randomly select a PM team for the case to be transferred to for cash/snap pending caseload
            team_to_check = all(i) & ""
            If left(team_to_check, len(team_to_check)-2) = left(caseload_type, len(caseload_type) - 2) Then
			    ReDim preserve temp_array(counter)
			    temp_array(counter) = things(i)
			    counter = counter + 1
		    End If
        Else
		    If all(i) = caseload_type Then
		    	ReDim preserve temp_array(counter)
		    	temp_array(counter) = things(i)
		    	counter = counter + 1
		    End If
        End If
	Next
	available_caseload_array = temp_array
end function

function gather_current_caseload(current_caseload, secondary_caseload, find_previous_pw, previous_pw)
	call navigate_to_MAXIS_screen("CASE", "CURR")
	EMReadScreen current_caseload, 7, 21, 14
	EMReadScreen secondary_caseload, 7, 21, 26
	current_caseload = trim(current_caseload)
	secondary_caseload = trim(secondary_caseload)
	If find_previous_pw = True Then
		call navigate_to_MAXIS_screen("SPEC", "XFER")
		call write_value_and_transmit("X", 5, 16)
		EMReadScreen previous_pw, 7, 18, 28
		PF3
		call navigate_to_MAXIS_screen("CASE", "CURR")
	End If

end function

app_facilities = "Monarch Facilities Contract"		'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at Bloomington "			'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at Chateau"					'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at Excelsior"				'MONARCH
app_facilities = app_facilities+chr(9)+"Estates at St. Louis Park"			'MONARCH
app_facilities = app_facilities+chr(9)+"Villas at Brookview"				'VILLAS 1
app_facilities = app_facilities+chr(9)+"Villas at The Park"				    'VILLAS 1
app_facilities = app_facilities+chr(9)+"Villas at Richfield"		        'VILLAS 1
app_facilities = app_facilities+chr(9)+"Villas at Robbinsdale"	            'VILLAS 1
app_facilities = app_facilities+chr(9)+"Villas at The Cedars"				'VILLAS 1
app_facilities = app_facilities+chr(9)+"Villas at Bryn Mawr"			    'VILLAS 2
app_facilities = app_facilities+chr(9)+"Villas at Osseo"					'VILLAS 2
app_facilities = app_facilities+chr(9)+"Villas at St. Louis Park"			'VILLAS 2
app_facilities = app_facilities+chr(9)+"Ebenezer Care Center/ Martin Luther Care Center"	'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Ebenezer Care Center"				'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Ebenezer Loren on Park"				'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Martin Luther Care Center"			'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"Meadow Woods"						'EBENEZER/MARTIN LUTHER
app_facilities = app_facilities+chr(9)+"North Memorial"
app_facilities = app_facilities+chr(9)+"HCMC"


'Initial dialog

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 241, 115, "Transfer Case"
  Text 5, 10, 50, 10, "Case Number:"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 90, 35, 45, 15, transfer_to_worker
  Text 5, 95, 60, 10, "Worker Signature:"
  EditBox 65, 90, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 90, 45, 15
    CancelButton 185, 90, 45, 15
  Text 10, 40, 75, 10, "New servicing worker:"
  GroupBox 5, 25, 220, 30, "Out of county only"
  GroupBox 5, 55, 220, 30, "Internal transfer only"
  Text 10, 70, 75, 10, "Reason for transfer:"
  DropListBox 80, 65, 140, 15, ""+chr(9)+"Change in circumstance"+chr(9)+"New application received"+chr(9)+"Case in wrong caseload/other", transfer_reason
EndDialog

'Runs the first dialog - which confirms the case number
DO
    active_worker_found = TRUE
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
			transfer_to_worker = trim(transfer_to_worker)

            Call validate_MAXIS_case_number(err_msg, "*")
            'IF len(transfer_to_worker) <> 7 THEN err_msg = err_msg & vbNewLine & "* Please enter the new servicing worker."
            IF UCASE(transfer_to_worker) = "X127CCL" then err_msg = err_msg & vbNewLine & "This case will be transferred via an automated script after being closed for 4 months. Choose another case load or press CANCEL to stop the script."
            IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF left(UCASE(trim(transfer_to_worker)), 4) = "X127" THEN err_msg = err_msg & vbNewLine & "For internal transfers, select internal transfer option and leave the worker ID as blank. This will ensure the case is transferred to the correct worker based on the programs on the case."
            If transfer_to_worker = "" and transfer_reason = "" then err_msg = err_msg & vbNewLine & "* Please select a reason for the transfer from the dropdown or enter a worker for out of county transfer."
            If transfer_to_worker <> "" and len(transfer_to_worker) <> 7 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid new servicing worker or clear the out of county field."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
    'TODO add a message / option to run app received if the reason for transfer is new application received

    transfer_to_worker = UCASE(trim(transfer_to_worker))         'reformatting to compare to other variables

    'Adding in a PRIV check and a background check/Script end if not passing background.
    Call back_to_SELF
    CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
    IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
    EMReadScreen SELF_check, 4, 2, 50
    If SELF_check = "SELF" then script_end_procedure_with_error_report("This case is still in background processing. Please complete any needed actions and rerun the script when the case has completed background processing.")

    If trim(transfer_to_worker) <> "" Then
        EMReadScreen worker_number, 7, 21, 17
        worker_number = UCASE(trim(worker_number))
        If transfer_to_worker = worker_number then script_end_procedure_with_error_report("This case is already in the worker number " & transfer_to_worker & ". The script will now end.")

        '----------Checks that the worker or agency is valid---------- 'must find user information before transferring to account for privileged cases.
        CALL navigate_to_MAXIS_screen("REPT", "USER")
        Call write_value_and_transmit(transfer_to_worker, 21, 12)
        'handling for inactive caseloads and for incorrect entry of a caseload. Will loop through the dialog again.
        EMReadScreen error_message, 75, 24, 2
        error_message = trim(error_message)
        EMReadScreen inactive_worker, 8, 7, 38
        IF inactive_worker = "INACTIVE" THEN
            active_worker_found = false
            err_msg = "* This worker does not appear to be active: " & transfer_to_worker & vbcr & "Enter a valid case load or x number."
        END IF
        IF error_message = "NO WORKER FOUND WITH THIS ID" THEN
            active_worker_found = false
            err_msg = "* No worker was found with this worker ID: " & transfer_to_worker & vbcr & "Enter a valid case load or x number."
        End if
        IF active_worker_found = FALSE THEN msgbox err_msg
    End If
LOOP UNTIL active_worker_found = TRUE

transfer_out_of_county = FALSE                                          'setting variable to false
IF transfer_to_worker <> "" Then
    transfer_out_of_county = TRUE 'setting the out of county BOOLEAN
Else
    '----------------------------------------------------------------------------------------------------In-county transfer
    Call gather_current_caseload(current_caseload, secondary_caseload, False, previous_pw)

    'Get the caseload type from current_caseload
    caseload_list = caseload_info.keys
    caseload_types = caseload_info.items
    For i = 0 to UBound(caseload_list)
        If caseload_list(i) = current_caseload Then
            current_caseload_type = caseload_types(i)
        End If
    Next

    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, active_programs, programs_applied_for)
    EMReadScreen pnd2_appl_date, 8, 8, 29               'Grabbing the PND2 date from CASE CURR in case the information cannot be pulled from REPT/PND2
      'collect member ages
      'collect expected caseload of the worker based on the programs on the case (function from application received?)
    Call back_to_SELF
    Call navigate_to_MAXIS_screen("STAT", "MEMB")
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadScreen age_of_memb_01, 3, 8, 76
    last_name = replace(last_name, "_", "")
    first_name = replace(first_name, "_", "")
    case_name = last_name & ", " & first_name
    age_of_memb_01 = trim(age_of_memb_01)
    If age_of_memb_01 = "" Then age_of_memb_01 = 0
    age_of_memb_01 = age_of_memb_01*1

    case_has_child_under_19 = False
    case_has_child_under_22 = False
    case_has_guardian = False
    Do
    	EMReadScreen memb_age, 3, 8, 76
    	memb_age = trim(memb_age)
    	If memb_age = "" Then memb_age = 0
    	memb_age = memb_age*1

    	If memb_age < 22 Then
    		case_has_child_under_22 = True
    		If memb_age < 19 Then case_has_child_under_19 = True
    	End If

    	EMReadScreen rel_to_applicant, 2, 10, 42
    	If rel_to_applicant = "03" Then case_has_guardian = True
    	If rel_to_applicant = "04" Then case_has_guardian = True
    	If rel_to_applicant = "18" Then case_has_guardian = True
    	If rel_to_applicant = "08" Then case_has_guardian = True
    	If rel_to_applicant = "09" Then case_has_guardian = True
    	If rel_to_applicant = "10" Then case_has_guardian = True
    	If rel_to_applicant = "11" Then case_has_guardian = True
    	If rel_to_applicant = "12" Then case_has_guardian = True
    	If rel_to_applicant = "13" Then case_has_guardian = True
    	If rel_to_applicant = "15" Then case_has_guardian = True
    	If rel_to_applicant = "16" Then case_has_guardian = True

    	transmit
    	EMReadScreen end_of_MEMB, 7, 24, 2
    Loop until end_of_MEMB = "ENTER A"

    preg_person_on_case = False
    Call navigate_to_MAXIS_screen("STAT", "PREG")
    Do
    	EMReadScreen preg_exists, 1, 2, 73
    	If preg_exists = "1" Then preg_person_on_case = True
    	transmit
    	EMReadScreen end_of_MEMB, 7, 24, 2
    Loop until end_of_MEMB = "ENTER A"

    Call navigate_to_MAXIS_screen("STAT", "PARE")
    EMReadScreen pare_exists, 1, 2, 73
    If pare_exists = "1" Then case_has_guardian = True


    child_under_19_question = "No"
    If case_has_child_under_19 = True Then child_under_19_question = "Yes"
    pregnant_question = "No"
    If preg_person_on_case = True Then pregnant_question = "Yes"

    'Main Dialog for in-county transfer cases - gathers information on the case to determine the correct caseload to transfer the case to. Information is gathered on the population of the case, healthcare programs on the case, and if the case is associated with a facility.
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 341, 210, "Case Info"
      ButtonGroup ButtonPressed
        OkButton 215, 185, 50, 15
        CancelButton 270, 185, 50, 15
      GroupBox 5, 5, 145, 100, "Population Info"
      GroupBox 165, 5, 160, 65, "HC Info"
      Text 10, 20, 75, 10, "Age of MEMB 01: " & age_of_memb_01
      Text 10, 30, 135, 20, "Case Name: " & case_name
      Text 10, 60, 100, 10, "Case has a member under 19:"
      Text 10, 75, 100, 10, "Case has a pregnant person:"
      DropListBox 115, 55, 30, 15, "No"+chr(9)+"Yes", child_under_19_question
      DropListBox 115, 70, 30, 15, "No"+chr(9)+"Yes", pregnant_question
      DropListBox 115, 85, 30, 15, "No"+chr(9)+"Yes", minor_parent_on_case
      Text 10, 155, 135, 10, "Current Caseload: " & current_caseload
      Text 10, 170, 135, 10, current_caseload_type
      GroupBox 165, 75, 160, 95, "Facility Info"
      Text 170, 135, 140, 10, "Resident Working with Contracted Facility:"
      DropListBox 170, 150, 140, 15, ""+chr(9)+app_facilities, contracted_facility
      CheckBox 170, 90, 150, 15, "Resides in / entering a HS facility.", hs_check
      DropListBox 215, 115, 95, 20, ""+chr(9)+"General"+chr(9)+"Long Term Homeless"+chr(9)+"HWS Facility"+chr(9)+"Demo Facility", faci_caseload_list
      Text 175, 120, 40, 10, "Faci Type:"
      GroupBox 5, 105, 145, 80, ""
      Text 10, 110, 55, 10, "Program Status:"
      Text 10, 125, 135, 10, "Active: " & active_programs
      Text 10, 140, 135, 10, "Pending: " & programs_applied_for
      DropListBox 170, 45, 145, 15, ""+chr(9)+"General Healthcare Pending"+chr(9)+"General Healthcare Active"+chr(9)+"LTC pending or requested"+chr(9)+"LTC Active"+chr(9)+"Waiver pending or requested"+chr(9)+"Waiver Active"+chr(9)+"TEFRA"+chr(9)+"MABC/SAGE"+chr(9)+"IV-E/Foster Care", HC_droplist
      Text 170, 20, 145, 20, "For cases with MAXIS healthcare, select general or specialty caseload types below:"
      Text 10, 90, 100, 10, "Case has parent(s) under 19:"

    EndDialog

    Do
        Do
            err_msg = ""
            dialog dialog1
            cancel_confirmation
            'TODO dialog errors
            'If the case has healthcare, ensure that a healthcare caseload type is selected. If case has GRH, ensure that the facility information is filled out. If case has a pending healthcare program, ensure that the status of the program is selected.
            If HC_droplist = "" and ((ma_status <> "INACTIVE" and ma_status <> "") or (msp_status <> "INACTIVE" and msp_status <> "")) AND (faci_caseload_list = "" AND contracted_facility = "") Then err_msg = err_msg & vbNewLine & "* Please select a healthcare caseload type from the dropdown."
            If (grh_status = "ACTIVE" OR grh_status = "PENDING" OR grh_status = "REIN" OR hs_check = 1) AND (faci_caseload_list = "" AND contracted_facility = "") Then err_msg = err_msg & vbNewLine & "* Please indicate if the resident is working with a facility and select the appropriate caseload type from the dropdown."
            If err_msg <> "" Then MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    loop until are_we_passworded_out = false

    'Logic to determine correct caseload based on the information gathered in the dialog and the information pulled from MAXIS on the case. This logic is based on the caseload directory that is maintained on the GitHub repository and is updated as needed when new caseloads are added or removed.
    If child_under_19_question = "Yes" Then
        minor_child_on_case = True
    Else
        minor_child_on_case = False
    End If

    If pregnant_question = "Yes" Then
        preg_person_on_case = True
    Else
        preg_person_on_case = False
    End If
    correct_caseload_type = ""
    'Specialty pops
    If hc_droplist = "TEFRA" Then correct_caseload_type = "TEFRA"
    If hc_droplist = "MABC/SAGE" Then correct_caseload_type = "MA - BC"
    If hc_droplist = "IV-E/Foster Care" Then correct_caseload_type = "Foster Care / IV-E"
    If hc_droplist = "MIPPA" Then correct_caseload_type = "MIPPA"

    'GRH/HS and 1800
    If correct_caseload_type = "" Then
		If grh_status = "ACTIVE" or grh_status = "PENDING" or grh_status = "REIN" or hs_check = checked Then
			correct_caseload_type = "GRH / HS"
			population = "Adults"
			If faci_caseload_list = "1800 - Team 160" Then
				correct_caseload_type = "1800 - Team 160"
				If InStr(alpha_split_one_a_l, left(case_name, 1)) <> 0 Then new_caseload = "X127EF8"
				If InStr(alpha_split_two_m_z, left(case_name, 1)) <> 0 Then new_caseload = "X127EF9"
				If new_caseload <> "" and current_caseload <> new_caseload Then transfer_needed = True
            ElseIf contracted_facility <> "" Then
                correct_caseload_type = "Contracted - " & contracted_facility
            Else
				If grh_status = "PENDING" Then
                    If faci_caseload_list = "Long Term Homeless" Then
                        correct_caseload_type = "GRH / HS - LTH Pending"
                    Else
                        correct_caseload_type = "GRH / HS - Adults Pending"
                    End If
                End If
                If grh_status = "ACTIVE" or grh_status = "REIN" Then
                    If faci_caseload_list = "General" Then correct_caseload_type = "GRH / HS - Maintenance"
                    If faci_caseload_list = "Long Term Homeless" Then correct_caseload_type = "GRH / HS - LTH Active"
                    If faci_caseload_list = "HWS Facility" Then correct_caseload_type = "GRH / HS - HWS"
                    If faci_caseload_list = "Demo Facility" Then correct_caseload_type = "GRH / HS - Demo"
                End If
				If minor_child_on_case = True or preg_person_on_case = True Then
					correct_caseload_type = "GRH / HS - Families Pending"
					population = "Families"
				End If
			End If
		End If
    End If
    'YET
    If correct_caseload_type = "" Then
        If age_of_memb_01 < 18 Then correct_caseload_type = "YET"
        If age_of_memb_01 >= 18 and age_of_memb_01 < 20 AND pregnant_question = "Yes" Then correct_caseload_type = "YET"
        If minor_parent_on_case = "Yes" Then correct_caseload_type = "YET"
    End If
    'LTC cases
    If correct_caseload_type = "" Then
        If HC_droplist = "LTC pending or requested" Then
            correct_caseload_type = "LTC+ - General"
            If InStr(alpha_split_one_a_h, left(case_name, 1)) <> 0 Then new_caseload = "X127EK4"
            If InStr(alpha_split_two_i_z, left(case_name, 1)) <> 0 Then new_caseload = "X127EK9"
        ElseIf HC_droplist = "LTC Active" Then
            correct_caseload_type = "LTC"
        End If
    End If
    'Waivers
    If correct_caseload_type = "" Then
        If HC_droplist = "Waiver pending or requested" Or HC_droplist = "Waiver Active" Then
           If child_under_19_question = "Yes" Then
                correct_caseload_type = "Waivers - Families"
            Else
                correct_caseload_type = "Waivers"
            End If
        End If
    End If
    Const GENERAL_HEALTHCARE_PENDING = "General Healthcare Pending"
    Const GENERAL_HEALTHCARE_ACTIVE = "General Healthcare Active"
    'General Health care
    If correct_caseload_type = "" Then
        If HC_droplist = GENERAL_HEALTHCARE_PENDING OR ma_status = "PENDING" or msp_status = "PENDING" Then
            correct_caseload_type = "Healthcare - Pending"
        End If
    End If

    If correct_caseload_type = "" Then
        If HC_droplist = GENERAL_HEALTHCARE_ACTIVE OR ma_status <> "INACTIVE" or msp_status  <> "INACTIVE" Then
            'TODO - split this into mixed cases with other programs vs. HC only
            If unknown_cash_pending = True or ga_status <> "INACTIVE" or msa_status <> "INACTIVE" or mfip_status <> "INACTIVE" or SNAP_status <> "INACTIVE" Then
                 correct_caseload_type = "Healthcare Mixed Active"
            Else
                correct_caseload_type = "Healthcare Only Active"
            End If
        End If
    End If
    'general maintenance
    If correct_caseload_type = "" Then
        If unknown_cash_pending = True or ga_status = "PENDING" or msa_status = "PENDING" or mfip_status = "PENDING" or SNAP_status = "PENDING" Then
            If child_under_19_question = "Yes" Then
                If mfip_status = "PENDING" OR unknown_cash_pending = True Then
                    correct_caseload_type = "Families - Cash"
                Else
                     correct_caseload_type = "Families - Pending"
                End If
            Else
                correct_caseload_type = "Adults - Pending"
            End IF
        Else
            If child_under_19_question = "Yes" Then
                correct_caseload_type = "Families Active"
            Else
                correct_caseload_type = "Adults Active"
            End If
        End If
    End If

    'Adding on the team to the caseload type for cases that are already on a program and therefore have a team in the current caseload type. This is to ensure that cases that are already on a program are transferred to the correct team for their caseload type.
    If left(correct_caseload_type, 6) = "Adults" OR left(correct_caseload_type, 8) = "Families"Then 'Grabs the current team for the caseload type for cases already active on a program
        If isnumeric(right(current_caseload_type, 1)) = True Then
            correct_caseload_type = correct_caseload_type & " " & right(current_caseload_type, 1)
        Else
            correct_caseload_type = correct_caseload_type & " ?" 'making correct length if current_caseload_type = ""
        End If
    End If
   'Error handling for trying to transfer to the same caseload type
    If correct_caseload_type = current_caseload_type Then
        no_transfer = MSgBox("The case is already in the correct caseload type based on the information provided. If you feel this is incorrect, press OK to proceed and enter the correct caseload and reason for transfer on the next dialog.", vbOkcancel)
        If no_transfer = vbCancel Then script_end_procedure("~PT: user pressed cancel")
    End If

    'Now select the right x# for transfer
    Call get_caseload_array_by_type(correct_caseload_type, available_caseload_array)
	number_of_options = UBound(available_caseload_array)
	Do
		pnd2_limit_issue = False
		call select_random_index(number_of_options, caseload_index)
		transfer_to_worker = available_caseload_array(caseload_index)
		If number_of_options = 0 Then Exit Do
		call navigate_to_MAXIS_screen("REPT", "PND2")
		Call write_value_and_transmit(transfer_to_worker, 21, 13)
        EMReadScreen pnd2_disp_limit, 13, 6, 35
        If pnd2_disp_limit = "Display Limit" Then
			pnd2_limit_issue = True
		Else
			EMReadScreen total_pages, 3, 3, 78
			total_pages = trim(total_pages)
			total_pages = total_pages * 1
			if total_pages > 195 Then pnd2_limit_issue = True
		End If
	Loop until pnd2_limit_issue = False

    'Transfer confirmation dialog - confirms the worker that the case is being transferred to and the caseload type the case is being transferred to. Also includes an option to report if the case is being transferred to a different worker than the one identified in the script with a reason why.
    dialog1 = ""
    BeginDialog Dialog1, 0, 0, 206, 135, "Dialog"
      ButtonGroup ButtonPressed
        OkButton 90, 110, 50, 15
        CancelButton 145, 110, 50, 15
      Text 5, 10, 180, 10, "This case will be transferred to the following caseload:"
      Text 5, 25, 175, 10, Correct_caseload_type
      Text 120, 25, 60, 10, transfer_to_worker
      GroupBox 5, 40, 190, 65, "Selected Caseload is incorrect"
      Text 10, 55, 55, 10, "Transfer to x#:"
      EditBox 65, 50, 60, 15, new_transfer_to_worker
      Text 10, 90, 30, 10, "Reason:"
      CheckBox 10, 70, 180, 10, "Check here if correct caseload cannot be determined", unknown_caseload_check
      EditBox 40, 85, 150, 15, error_reason
    EndDialog


    Do
        err_msg = ""
        dialog dialog1
        cancel_confirmation
        if (new_transfer_to_worker <> "" OR unknown_caseload_check = checked) and error_reason = "" Then err_msg = err_msg & vbcr & "* Please provide a reason for transferring to a different worker than the one identified in the dialog."
        If err_msg <> "" Then MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""


    If new_transfer_to_worker <> ""  OR unknown_caseload_check = checked Then
         bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
        subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"

        full_text = "Error occurred on " & date & " at " & time
        full_text = full_text & vbCr & "The case transfer for case " & MAXIS_case_number & " was transferred to a different worker than the one identified in the script."
        full_text = full_text & vbCr & "Reason for different worker transfer: " & error_reason
        full_text = full_text & vbCr & "Original worker identified: " & transfer_to_worker & vbCrLf & "New worker entered: " & new_transfer_to_worker

        Call create_outlook_email("", bzt_email, "", "", subject_of_email, 1, False, "", "", False, "", full_text, True, "", True)
        transfer_to_worker = new_transfer_to_worker
        're-align the caseload type with the new_transfer_to_worker to keep the stats correct

        caseload_list = caseload_info.keys
        caseload_types = caseload_info.items
        For i = 0 to UBound(caseload_list)
            If caseload_list(i) = transfer_to_worker Then
                correct_caseload_type = caseload_types(i)
            End If
        Next
    End If

    If unknown_caseload_check = checked Then
        script_end_procedure_with_error_report("The correct caseload for this case could not be determined. Please review the case and transfer to the correct caseload as needed. Reason provided: " & error_reason)
    Else
        transfer_to_worker = UCASE(trim(transfer_to_worker))
    End If
    'Staff need to confirm that they've transferred the case in ECF Next 1st
    Do
        Do
            transfer_confirmation = MsgBox("Has this case been transferred in ECF Next? New worker: " & transfer_to_worker, vbQuestion + vbYesNo, "Transfer Case Reminder")
            If transfer_confirmation = vbNo then msgbox "This case needs to be transferred in ECF Next first." & vbcr & vbcr & "Transfer the case in ECF Next now."
            If transfer_confirmation = vbYes then exit do
        Loop
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER IN COUNTY
    Call write_value_and_transmit("X", 7, 16)              'transfer within county option
    PF9                                                    'putting the transfer in edit mode
    EMreadscreen second_servicing_worker, 7, 18, 74        'checking to see if the transfer_to_worker is the same as the second_servicing_worker (because then it won't transfer)
    IF second_servicing_worker <> "_______" THEN CALL clear_line_of_text(18, 74)
    Call write_value_and_transmit(transfer_to_worker, 18, 61)           'entering the worker information & saving - this should then take us to the transfer menu
    EMReadScreen servicing_worker, 7, 24, 30  'if it is not the transfer_to_worker - the transfer failed.
    EMReadScreen transfer_message, 79, 24, 2
    transfer_message = trim(transfer_message)

    STATS_manualtime = 120                      'manual run time in seconds
    'Script low down for inner county transfer
    script_run_lowdown = script_run_lowdown & "transfer_out_of_county: " & transfer_out_of_county & vbCr & "worker_number: " & worker_number & vbCr & "transfer_to_worker: " & transfer_to_worker & vbCr & " Error Message at transfer: " & vbCr & error_message
    If servicing_worker <> transfer_to_worker THEN
		end_msg = "Transfer of this case to " & transfer_to_worker & " has failed."
		end_msg = end_msg & vbCr & vbCr & "Message from MAXIS on XFER screen:" & vbCr & transfer_message
		script_end_procedure_with_error_report("Transfer of this case to " & transfer_to_worker & " has failed.")
	End If
	end_msg = "Success! Case transferred from: " & current_caseload & "|" & current_caseload_type & " to: " & transfer_to_worker & "|" & correct_caseload_type & " Due to: " & transfer_reason

    script_end_procedure_with_error_report(end_msg)
End if

'----------------------------------------------------------------------------------------------------Out-of-County Case Transfer
Call write_value_and_transmit("X", 7, 3) ' navigating to read the worker information from REPT/USER
EMReadScreen worker_name, 43, 8, 27       'reading the worker ALIAS name 1st
worker_name = trim(worker_name)

IF worker_name = "" THEN 						'If we are unable to find the alias for the worker we will just use the worker name as it is what is used on notices anyway
	EMReadScreen worker_name, 43, 7, 27         'reading the worker name if the ALIAS is ""
	worker_name = trim(worker_name)
	name_length = len(worker_name)
	comma_location = InStr(worker_name, ",")
	worker_name = right(worker_name, (name_length - comma_location)) & " " & left(worker_name, (comma_location - 1)) 'this section will reorder the name of the worker since it is stored here as last, first. the comma_location - 1 removes the comma from the "last,"
END IF

EMReadScreen mail_addr_line_one, 43, 9, 27              ' really only need for out of county but read for all '
	mail_addr_line_one = trim(mail_addr_line_one)
EMReadScreen mail_addr_line_two, 43, 10, 27
	mail_addr_line_two = trim(mail_addr_line_two)
EMReadScreen mail_addr_line_three, 43, 11, 27
	mail_addr_line_three = trim(mail_addr_line_three)
EMReadScreen mail_addr_line_four, 43, 12, 27
	mail_addr_line_four = trim(mail_addr_line_four)
EMReadScreen worker_agency_phone, 14, 13, 27
EMReadScreen worker_county_code, 2, 15, 32

'MNPrairie Bank Support - MNPrairie Bank cases all go to Steele (county code 74)'s ICT transfer.
'Agencies in the MNPrairie Bank are Dodge (county code 20), Steele (county code 74), and Waseca (county code 81)
IF transfer_to_worker = "X120ICT" OR transfer_to_worker = "X181ICT" THEN transfer_to_worker = "X174ICT"
IF transfer_to_worker = "X162ICT" THEN worker_agency_phone = "651-266-4444" 'Ramsey County has an individuals workers phone previously'

CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

cash_cfr = "27" 'defaulting to Hennepin for out of county
hc_cfr   = "27"

IF grh_status = "ACTIVE" or grh_status = "APP OPEN" then excluded_time_dropdown = "Yes" 'If GRH is Pending, REIN or closing then excluded time wouldn't apply
'-------------------------------------------------------------------------------------------------DIALOG

Do
    Do
        'Inside the do...loop for the calculate_button
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 346, 260, "Out-of-County Case Transfer for #" & MAXIS_case_number
            Text 10, 10, 70, 10, "Resident Move Date:"
            EditBox 80, 5, 45, 15, resident_move_date
            Text 135, 10, 55, 10, "Excluded time?"
            DropListBox 190, 5, 45, 15, "Select:"+chr(9)+"No"+chr(9)+"Yes", excluded_time_dropdown
            Text 245, 10, 40, 10, "Begin Date:"
            EditBox 285, 5, 45, 15, excluded_time_begin_date
            Text 30, 35, 45, 10, "METS Status:"
            DropListBox 80, 30, 45, 15, "Select:"+chr(9)+"active"+chr(9)+"inactive"+chr(9)+"pending"+chr(9)+"N/A", mets_status_dropdown
            Text 140, 35, 50, 10, "METS Case #:"
            EditBox 190, 30, 45, 15, METS_case_number
            Text 5, 60, 70, 10, "Reason For Transfer:"
            EditBox 80, 55, 260, 15, transfer_reason
            Text 15, 80, 60, 10, "Outstanding Work:"
            EditBox 80, 75, 260, 15, outstanding_work
            Text 5, 100, 135, 10, "List All Requested/Pending Verifications:"
            EditBox 140, 95, 200, 15, requested_verifs
            Text 5, 120, 195, 10, "Note any expected changes in household's circumstances:"
            EditBox 200, 115, 140, 15, expected_changes
            Text 5, 140, 45, 10, "Other Notes:"
            EditBox 50, 135, 290, 15, other_notes
            GroupBox 5, 155, 335, 60, "County of Financial Responsibility(CFR) - Complete if not excluded time."
            Text 15, 175, 100, 10, "Cash Programs  - Current CFR:"
            EditBox 115, 170, 20, 15, cash_cfr
            Text 25, 195, 85, 10, "Health Care - Current CFR:"
            EditBox 115, 190, 20, 15, hc_cfr
            Text 160, 175, 75, 10, "Change Date (MM YY):"
            EditBox 235, 170, 20, 15, cash_cfr_month
            EditBox 260, 170, 20, 15, cash_cfr_year
            Text 160, 195, 75, 10, "Change Date (MM YY):"
            EditBox 235, 190, 20, 15, hc_cfr_month
            EditBox 260, 190, 20, 15, hc_cfr_year
            ButtonGroup ButtonPressed
                PushButton 290, 175, 40, 25, "Calculate", calculate_button
                OkButton 235, 235, 50, 15
                CancelButton 290, 235, 50, 15
            GroupBox 5, 220, 220, 35, "Navigation:"
            ButtonGroup ButtonPressed
                PushButton 15, 235, 50, 15, "TE02.08.133", POLI_TEMP_button
                PushButton 65, 235, 50, 15, "SPEC/XFER", XFER_button
                PushButton 115, 235, 50, 15, "Use Form", useform_xfer_button
                PushButton 165, 235, 50, 15, "HSRM-XFER", hsrm_xfer_button
                           Text 10, 285, 45, 10, "Current CFR:"
            Text 85, 285, 75, 10, "Change Date (MM YY)"
            ButtonGroup ButtonPressed
        EndDialog

		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
        IF IsDate(resident_move_date) = False OR Len(resident_move_date) <> 10 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date in the MM/DD/YYYY format for resident move."
        If excluded_time_dropdown = "Select:" then err_msg = err_msg & vbNewLine & "* Indicate if this is an excluded time case."
        IF excluded_time_dropdown = "Yes" then
            If IsDate(excluded_time_begin_date) = False OR Len(excluded_time_begin_date) <> 10 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date in the MM/DD/YYYY format for the start of excluded time or double check that the resident's absence is due to excluded time."
        End if
        If mets_status_dropdown = "Select:" then err_msg = err_msg & vbNewLine & "* Select a METS status."
		IF mets_status_dropdown = "active" or mets_status_dropdown = "pending" then
            If IsNumeric(METS_case_number) = False or len(METS_case_number) <> 8 then err_msg = err_msg & vbNewLine & "* Please enter a valid 8-digit METS case number."
        End if
        IF trim(transfer_reason) = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a reason for transfer."
        If case_pending = True and trim(requested_verifs) = "" then err_msg = err_msg & vbNewLine & "* List the verifications requested that are still needed for the program(s) : " & list_pending_programs & "."

        If excluded_time_dropdown = "No" then
            IF (ga_case = TRUE or msa_case = TRUE or mfip_case = TRUE or dwp_case = TRUE) THEN
                If isnumeric(cash_cfr) = False or len(cash_cfr) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit county code Current Financial Responsibility County (CFR) code for cash."
                If isnumeric(cash_cfr_month) = False or len(cash_cfr_month) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit month for the Current Financial Responsibility County (CFR) code for cash."
                If isnumeric(cash_cfr_year) = False or len(cash_cfr_year) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit year for the Current Financial Responsibility County (CFR) code for cash."
            End if
            If ma_case = True THEN
                If isnumeric(hc_cfr) = False or len(hc_cfr) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit county code Current Financial Responsibility County (CFR) code for Health Care."
                If isnumeric(hc_cfr_month) = False or len(hc_cfr_month) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit month for the Current Financial Responsibility County (CFR) code for Health Care."
                If isnumeric(hc_cfr_year) = False or len(hc_cfr_year) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit year for the Current Financial Responsibility County (CFR) code for Health Care."
            End if
        End if

        IF ButtonPressed = POLI_TEMP_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.08.133%20COMPLETING%20AN%20INTER-AGENCY%20CASE%20TRANSFER.pdf?csf=1&web=1&e=4jRZFD"
        IF ButtonPressed = XFER_button THEN CALL MAXIS_dialog_navigation()
        IF ButtonPressed = useform_xfer_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://aem.hennepin.us/rest/services/HennepinCounty/Processes/ServletRenderForm:1.0?formName=HSPH5069_1-0.xdp&interactive=1"
        IF ButtonPressed = hsrm_xfer_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Case-Transfers.aspx"
        If ButtonPressed = calculate_button then
            'Determining the CFR based on the date of move.
            CM_plus_3_mo =  right("0" &  DatePart("m",    DateAdd("m", 3, date)), 2)
            CM_plus_3_yr =  right(       DatePart("yyyy", DateAdd("m", 3, date)), 2)

            If datepart("d", resident_move_date ) = "1" then
                cash_cfr_month = CM_plus_2_mo
                cash_cfr_year = CM_plus_2_yr
            Else
                cash_cfr_month = CM_plus_3_mo
                cash_cfr_year = CM_plus_3_yr
            End if

            hc_cfr_month = cash_cfr_month
            hc_cfr_year = cash_cfr_year
        End if

        IF ButtonPressed = useform_xfer_button or ButtonPressed = XFER_button or ButtonPressed = POLI_TEMP_button or ButtonPressed = calculate_button or ButtonPressed = hsrm_xfer_button THEN
            err_msg = "LOOP"
        Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        End If
    Loop until err_msg = ""
    IF ButtonPressed = useform_xfer_button or ButtonPressed = XFER_button or ButtonPressed = POLI_TEMP_button or ButtonPressed = hsrm_xfer_button THEN call back_to_self ' this is for if the worker has used the POLI/TEMP navigation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = False					'loops until user passwords back in

'SENDING a SPEC/MEMO - this happens before the case note, and transfer -  we overwrite the information
'----------Sending the resident a SPEC/MEMO notifying them of the details of the transfer----------
Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    		'Writes the appt letter into the MEMO.
Call write_variable_in_SPEC_MEMO("Your case has been transferred. Your new agency/worker is: " & worker_name & "")
Call write_variable_in_SPEC_MEMO("If you have any questions, or to send in requested proofs,")
Call write_variable_in_SPEC_MEMO("please direct all communications to the agency listed.")
Call write_variable_in_SPEC_MEMO(worker_name)
Call write_variable_in_SPEC_MEMO(mail_addr_line_one)
Call write_variable_in_SPEC_MEMO(mail_addr_line_two)
Call write_variable_in_SPEC_MEMO(mail_addr_line_three)
Call write_variable_in_SPEC_MEMO(mail_addr_line_four)
Call write_variable_in_SPEC_MEMO(worker_agency_phone)
Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
Call write_variable_in_SPEC_MEMO("You can also request a paper copy. Auth: 7CFR 273.2(e)(3).")
PF4'save and exit

'----------The case note of the reason for the XFER----------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("~ Case transferred to " & transfer_to_worker & " ~")
call write_bullet_and_variable_in_case_note("Resident move date", resident_move_date)
call write_bullet_and_variable_in_case_note("Change report sent", date) 'defaulting information '
call write_bullet_and_variable_in_case_note("Case file sent:", date) 'defaulting information '
Call write_bullet_and_variable_in_CASE_NOTE("Reason case was transferred", transfer_reason)
Call write_variable_in_CASE_NOTE("--")
Call write_bullet_and_variable_in_CASE_NOTE("METS Status", mets_status_dropdown)
If mets_status_dropdown = "active" or mets_status_dropdown = "pending" then Call write_bullet_and_variable_in_CASE_NOTE("METS Case Number", METS_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("Active programs", list_active_programs)
Call write_bullet_and_variable_in_CASE_NOTE("Pending programs", list_pending_programs)
Call write_variable_in_CASE_NOTE("--")
Call write_bullet_and_variable_in_CASE_NOTE("Outstanding case work", outstanding_work)
Call write_bullet_and_variable_in_CASE_NOTE("Requested verifications", requested_verifs)
Call write_bullet_and_variable_in_CASE_NOTE("Expected changes in household's circumstances:", expected_changes)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("--")
IF excluded_time_dropdown = "Yes" THEN
    call write_bullet_and_variable_in_case_note("Excluded Time" , "Yes, Begins " & excluded_time_begin_date)
ELSEIF excluded_time_dropdown = "No" THEN
    call write_bullet_and_variable_in_case_note("Excluded Time", excluded_time_dropdown)
    IF ma_case = True or msp_case = True THEN
        CALL write_bullet_and_variable_in_case_note("HC County of Financial Responsibility", hc_cfr)
        CALL write_bullet_and_variable_in_case_note("HC CFR Change Date", (hc_cfr_month & "/" & hc_cfr_year))
    END IF
    IF (ga_case = TRUE or msa_case = TRUE or mfip_case = TRUE or dwp_case = TRUE or grh_case = TRUE) THEN
        CALL write_bullet_and_variable_in_case_note("CASH County of Financial Responsibility", cash_cfr) 'county_financial_responsibility'
        CALL write_bullet_and_variable_in_case_note("CASH CFR Change Date", (cash_cfr_month & "/" & cash_cfr_year))
    END IF
    Call write_variable_in_CASE_NOTE("--")
END IF
CALL write_variable_in_case_note("* SPEC/MEMO sent to resident with new worker information.")
IF forms_to_arep = "Y" THEN call write_variable_in_case_note("* Copy of SPEC/MEMO sent to AREP.")
IF forms_to_swkr = "Y" THEN call write_variable_in_case_note("* Copy of SPEC/MEMO sent to social worker.")
Call write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE (worker_signature)
PF3

'----------------------------------------------------------OUT OF COUNTY TRANSFER actually happening
CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER
Call write_value_and_transmit("X", 9, 16)                               'Transfer County To County
EMReadScreen panel_check, 4, 2, 54                        'reading to see if we made it to the right place
IF panel_check = "XCTY" THEN
    EMReadScreen low_down_Excluded_Time_Begin_Date, 8, 6, 28    'reading for script script_run_lowdown'
    EMReadScreen Cash_I,  2, 11, 39                              'reading for script script_run_lowdown'
    EMReadScreen Cash_II, 2, 12, 39                              'reading for script script_run_lowdown'
    EMReadScreen GRH, 2, 13, 39                                  'reading for script script_run_lowdown'
    EMReadScreen Health_Care, 2, 14, 39                          'reading for script script_run_lowdown'
    EMReadScreen MA_Excluded_Time, 2, 15, 39                     'reading for script script_run_lowdown'
    EMReadScreen IVE_Foster_Care, 2, 16, 39                     'reading for script script_run_lowdown'
    PF9                                                          'putting the transfer in edit mode

    IF low_down_Excluded_Time_Begin_Date <> "__ __ __" then
        CALL clear_line_of_text(6, 28)  'clearing any old Excluded Time Begin Date __ __ __'
        CALL clear_line_of_text(6, 31)
        CALL clear_line_of_text(6, 34)
    End if

    If Cash_I <> "__" then CALL clear_line_of_text(11, 39) 'Clearing Current Fin Resp County Cash I'
    If Cash_II <> "__" then CALL clear_line_of_text(12, 39) 'Cash II'
    If Health_Care <> "__" then CALL clear_line_of_text(14, 39) 'Health Care'
    If MA_Excluded_Time <> "__" then CALL clear_line_of_text(15, 39) 'MA Excluded Time'
    IF IVE_Foster_Care <> "__" then CALL clear_line_of_text(16, 39) 'IV-E Foster Care

    call create_MAXIS_friendly_date(resident_move_date, 0, 4, 28)    'Writing resident move date
    call create_MAXIS_friendly_date(date, 0, 4, 61)                  'this is the Change Report Sent date we dont need to ask because we dont do this'
    call create_MAXIS_friendly_date(date, 0, 5, 61)                  'this is the Case File Sent date we dont need to ask because we dont do this'

    EMWriteScreen left(excluded_time_dropdown, 1), 5, 28            'Writes the excluded time info. Only need the left character (it's a dropdown)

    IF excluded_time_dropdown = "Yes" THEN                          'If there's excluded time, need to write the info
        call create_MAXIS_friendly_date(excluded_time_begin_date, 0, 6, 28)
        EMWriteScreen hc_cfr, 15, 39
    Else
        IF (ga_case = TRUE or msa_case = TRUE or mfip_case = TRUE or dwp_case = TRUE) THEN
            EmReadscreen Cash_I_prog, 2, 11, 28
            If trim(Cash_I_prog) <> "" then
                EMWriteScreen cash_cfr, 11, 39
                EMWriteScreen cash_cfr_month, 11, 53
                EMWriteScreen cash_cfr_year, 11, 59
            End if
            EmReadscreen Cash_I_prog, 2, 12, 28
            If trim(Cash_II_prog) <> "" then
                EMWriteScreen cash_cfr, 12, 39
                EMWriteScreen cash_cfr_month, 12, 53
                EMWriteScreen cash_cfr_year, 12, 59
            End if
	        EMWriteScreen cash_cfr_year, 11, 59
            EMWriteScreen cash_cfr, 12, 39 'cash II because I blank it out there is no need to read'
        END IF

        If grh_case = TRUE then EMWriteScreen cash_cfr, 13, 39 'for GRH'

        IF ma_case = TRUE or msp_case = True THEN
            EMWriteScreen hc_cfr, 14, 39
            EMWriteScreen hc_cfr_month, 14, 53
            EMWriteScreen hc_cfr_year, 14, 59
        END IF
    End if

    EMreadscreen second_servicing_worker, 7, 18, 74        'checking to see if the transfer_to_worker is the same as the second_servicing_worker (because then it won't transfer)
    IF second_servicing_worker <> "_______" THEN CALL clear_line_of_text(18, 74)
    Call write_value_and_transmit(transfer_to_worker, 18, 61)           'entering the worker information & saving - this should then take us to the transfer menu
    EMReadScreen servicing_worker, 7, 24, 30  'if it is not the transfer_to_worker - the transfer failed.
End if

STATS_manualtime = 300              'manual run time in seconds
'Script low down for intercounty transfer

script_run_lowdown = script_run_lowdown & "transfer_out_of_county: " & transfer_out_of_county & vbCr & "excluded_time_dropdown: " & excluded_time_dropdown & vbCr & "Resident_move_date: " & Resident_move_date & vbcr & _
"Excluded_time_begin_date: " & Excluded_time_begin_date & vbcr & "mets_status_dropdown: " & mets_status_dropdown & vbcr & "mets_case_number:" & mets_case_number & vbcr & "transfer_reason: " & transfer_reason & Vbcr & _
"case_pending: " & case_pending & vbcr & "ga_case: " & ga_case & vbCr & "msa_case: " & msa_case & vbCr & "dwp_case: " & dwp_case & vbCr & "grh_case: " & grh_case & vbCr & "ma_case: " & ma_case & vbCr & "msp_case: " & msp_case & vbCr & _
"outstanding_work: " & outstanding_work & vbCr & "requested_verifs: " & requested_verifs & vbCr & "expected_changes: " & expected_changes & vbCr & "other_notes: " & other_notes & vbCr & _
"cash cfr county code: " & cash_cfr & vbCr & "cash cfr date: " & cash_cfr_month & "/" & cash_cfr_year & vbcr & "hc_cfr county code: " & hc_cfr & vbcr & "HC cfr date: " & hc_cfr_month & "/" & hc_cfr_year & vbCr & _
"worker_number: " & worker_number & vbCr & "transfer_to_worker: " & transfer_to_worker & vbCr & " Error Message at transfer: " & vbCr & error_message

'TODO - set up end_msg and script run lowdown for data for internal transfer

If servicing_worker <> transfer_to_worker THEN
    script_end_procedure_with_error_report("Transfer of this case to " & transfer_to_worker & " has failed.")
Else
    script_end_procedure_with_error_report("Case transfer has been completed to: " & transfer_to_worker & ".")
End if

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/26/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/23/2022
'--Mandatory fields all present & Reviewed--------------------------------------10/26/2022
'--All variables in dialog match mandatory fields-------------------------------10/26/2022
'Review dialog names for content and content fit in dialog----------------------01/12/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/26/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------10/26/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------10/26/2022
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-10/26/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/26/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------10/26/2022
'--PRIV Case handling reviewed -------------------------------------------------10/26/2022
'--Out-of-County handling reviewed----------------------------------------------10/26/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/26/2022
'--BULK - review output of statistics and run time/count (if applicable)--------10/26/2022------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------10/26/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/26/2022
'--Incrementors reviewed (if necessary)-----------------------------------------10/26/2022------------------N/A
'--Denomination reviewed -------------------------------------------------------10/26/2022
'--Script name reviewed---------------------------------------------------------10/26/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------10/26/2022------------------N/A
'
'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------01/23/2024
'--comment Code-----------------------------------------------------------------01/23/2024
'--Update Changelog for release/update------------------------------------------01/23/2024
'--Remove testing message boxes-------------------------------------------------01/23/2024
'--Remove testing code/unnecessary code-----------------------------------------01/23/2024
'--Review/update SharePoint instructions----------------------------------------01/23/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------01/23/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------01/23/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------01/23/2024
'--Complete misc. documentation (if applicable)---------------------------------01/23/2024
'--Update project team/issue contact (if applicable)----------------------------01/23/2024
