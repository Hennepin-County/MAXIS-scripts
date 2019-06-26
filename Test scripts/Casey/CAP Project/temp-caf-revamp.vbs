
'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CAF.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 720                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

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
Call changelog_update("04/10/2019", "There was a bug that sometimes made the dialogs write over each other and be illegible, updated the script to keep this from happening.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
CALL changelog_update("10/17/2018", "Updated dialog box to reflect currecnt application process.", "MiKayla Handley, Hennepin County")
call changelog_update("05/05/2018", "Added autofill functionality for the DIET panel. NAT errors have been resolved.", "Ilse Ferris, Hennepin County")
call changelog_update("05/04/2018", "Removed autofill functionality for the DIET panel temporarily until MAXIS help desk can resolve NAT errors.", "Ilse Ferris, Hennepin County")
call changelog_update("01/11/2017", "Adding functionality to offer a TIKL for 12 month contact on 24 month SNAP renewals.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================
function HH_comp_dialog(HH_member_array)
'--- This function creates an array of all household members in a MAXIS case, and allows users to select which members to seek/add information to add to edit boxes in dialogs.
'~~~~~ HH_member_array: should be HH_member_array for function to work
'===== Keywords: MAXIS, member, array, dialog
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

    member_count = 0
    adult_cash_count = 0
    child_cash_count = 0
    adult_snap_count = 0
    child_snap_count = 0
    adult_emer_count = 0
    child_emer_count = 0
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadscreen last_name, 25, 6, 30
		EMReadscreen first_name, 12, 6, 63
		EMReadscreen mid_initial, 1, 6, 79
        EMReadScreen memb_age, 3, 8, 76
        memb_age = trim(memb_age)
        If memb_age = "" Then memb_age = 0
        memb_age = memb_age * 1
		last_name = trim(replace(last_name, "_", ""))
		first_name = trim(replace(first_name, "_", ""))
		mid_initial = replace(mid_initial, "_", "")

        ReDim Preserve ALL_MEMBERS_ARRAY(clt_notes, member_count)

        ALL_MEMBERS_ARRAY(memb_numb, member_count) = ref_nbr
        ALL_MEMBERS_ARRAY(clt_name, member_count) = last_name & ", " & first_name & " " & mid_initial

        If cash_checkbox = checked Then
            ALL_MEMBERS_ARRAY(include_cash_checkbox, member_count) = checked
            ALL_MEMBERS_ARRAY(count_cash_checkbox, member_count) = checked
            If memb_age > 18 then
                adult_cash_count = adult_cash_count + 1
            Else
                child_cash_count = child_cash_count + 1
            End If
        End If
        If SNAP_checkbox = checked Then
            ALL_MEMBERS_ARRAY(include_snap_checkbox, member_count) = checked
            ALL_MEMBERS_ARRAY(count_snap_checkbox, member_count) = checked
            If memb_age > 21 then
                adult_snap_count = adult_snap_count + 1
            Else
                child_snap_count = child_snap_count + 1
            End If
        End If
        If EMER_checkbox = checked Then
            ALL_MEMBERS_ARRAY(include_emer_checkbox, member_count) = checked
            ALL_MEMBERS_ARRAY(count_emer_checkbox, member_count) = checked
            If memb_age > 18 then
                adult_emer_count = adult_emer_count + 1
            Else
                child_emer_count = child_emer_count + 1
            End If
        End If

		client_string = ref_nbr & last_name & first_name & mid_initial
		client_array = client_array & client_string & "|"
		transmit
		Emreadscreen edit_check, 7, 24, 2
        member_count = member_count + 1
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

    Call navigate_to_MAXIS_screen("STAT", "PARE")
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
        transmit

        EMReadScreen panel_check, 14, 24, 13
        If panel_check <> "DOES NOT EXIST" Then
            pare_row = 8
            Do
                EMReadScreen child_ref_nbr, 2, pare_row, 24
                EMReadScreen rela_type, 1, pare_row, 53
                EMReadScreen rela_verif, 2, pare_row, 71

                If rela_type = "1" then relationship_type = "Parent"
                If rela_type = "2" then relationship_type = "Stepparent"
                If rela_type = "3" then relationship_type = "Grandparent"
                If rela_type = "4" then relationship_type = "Relative Caregiver"
                If rela_type = "5" then relationship_type = "Foster parent"
                If rela_type = "6" then relationship_type = "Caregiver"
                If rela_type = "7" then relationship_type = "Guardian"
                If rela_type = "8" then relationship_type = "Relative"

                If rela_verif = "BC" Then relationship_verif = "Birth Certificate"
                If rela_verif = "AR" Then relationship_verif = "Adoption Records"
                If rela_verif = "LG" Then relationship_verif = "Legal Guardian"
                If rela_verif = "RE" Then relationship_verif = "Religious Records"
                If rela_verif = "HR" Then relationship_verif = "Hospital Records"
                If rela_verif = "RP" Then relationship_verif = "Recognition of Parantage"
                If rela_verif = "OT" Then relationship_verif = "Other"
                If rela_verif = "NO" Then relationship_verif = "NONE"

                If child_ref_nbr <> "__" Then relationship_detail = relationship_detail & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " is the " & relationship_type & " of Memb " & child_ref_nbr & " - Verif: " & relationship_verif & "; "
                pare_row = pare_row + 1
            Loop Until child_ref_nbr = "__"
        End If
    Next

    If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", EATS)

	' client_array = TRIM(client_array)
	' test_array = split(client_array, "|")
	' total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
    '
	' DIM all_client_array()
	' ReDim all_clients_array(total_clients, 1)
    '
	' FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	' 	Interim_array = split(client_array, "|")
	' 	all_clients_array(x, 0) = Interim_array(x)
	' 	all_clients_array(x, 1) = 1
	' NEXT




    Do
        Do
            err_msg = ""
            adult_cash_count = adult_cash_count & ""
            child_cash_count = child_cash_count & ""
            adult_snap_count = adult_snap_count & ""
            child_snap_count = child_snap_count & ""
            adult_emer_count = adult_emer_count & ""
            child_emer_count = child_emer_count & ""

            dlg_len = 115 + (15 * UBound(ALL_MEMBERS_ARRAY, 2))
            if dlg_len < 130 Then dlg_len = 130
            BeginDialog Dialog1, 0, 0, 446, dlg_len, "HH Composition Dialog"
              Text 10, 10, 250, 10, "This dialog will clarify the household relationships and details for the case."
              Text 105, 25, 100, 10, "Included and Counted  in Grant"
              Text 110, 40, 20, 10, "Cash"
              Text 145, 40, 20, 10, "SNAP"
              Text 180, 40, 20, 10, "EMER"
              Text 230, 25, 90, 10, "Income Counted - Deeming"
              Text 230, 40, 20, 10, "Cash"
              Text 265, 40, 20, 10, "SNAP"
              Text 300, 40, 20, 10, "EMER"
              GroupBox 330, 5, 110, 100, "HH Count by program"
              Text 335, 15, 100, 20, "Enter the number of adults and children for each program"
              Text 370, 35, 20, 10, "Adults"
              Text 400, 35, 30, 10, "Children"
              Text 345, 50, 20, 10, "Cash"
              EditBox 370, 45, 20, 15, adult_cash_count
              EditBox 405, 45, 20, 15, child_cash_count
              Text 345, 70, 20, 10, "SNAP"
              EditBox 370, 65, 20, 15, adult_snap_count
              EditBox 405, 65, 20, 15, child_snap_count
              Text 345, 90, 25, 10, "EMER"
              EditBox 370, 85, 20, 15, adult_emer_count
              EditBox 405, 85, 20, 15, child_emer_count
              y_pos = 55
              For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                  Text 10, y_pos, 100, 10, ALL_MEMBERS_ARRAY(clt_name, each_member)
                  CheckBox 115, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member)
                  CheckBox 150, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member)
                  CheckBox 185, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member)
                  CheckBox 235, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member)
                  CheckBox 270, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member)
                  CheckBox 305, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member)
                  y_pos = y_pos + 15
              Next
              y_pos = y_pos + 5
              Text 10, y_pos + 5, 25, 10, "EATS:"
              EditBox 35, y_pos, 290, 15, EATS
              Text 10, y_pos + 25, 90, 10, "Household Relationships:"
              EditBox 105, y_pos + 20, 220, 15, relationship_detail
              ButtonGroup ButtonPressed
                OkButton 335, y_pos + 20, 50, 15
                CancelButton 390, y_pos + 20, 50, 15
            EndDialog

            dialog Dialog1
            cancel_without_confirmation

            adult_cash_count = adult_cash_count * 1
            child_cash_count = child_cash_count * 1
            adult_snap_count = adult_snap_count * 1
            child_snap_count = child_snap_count * 1
            adult_emer_count = adult_emer_count * 1
            child_emer_count = child_emer_count * 1

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    HH_member_array = ""

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member) = checked Then HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        If ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member) = checked Then HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        If ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member) = checked Then HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        If ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member) = checked Then HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
        If ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member) = checked Then HH_member_array = HH_member_array & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " "
    Next
	' BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	' 	Text 10, 5, 105, 10, "Household members to look at:"
	' 	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
	' 		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	' 	NEXT
	' 	ButtonGroup ButtonPressed
	' 	OkButton 185, 10, 50, 15
	' 	CancelButton 185, 30, 50, 15
	' ENDDIALOG
	' 												'runs the dialog that has been dynamically created. Streamlined with new functions.
	' Dialog HH_memb_dialog
	' If buttonpressed = 0 then stopscript
	' check_for_maxis(True)
    '
	' HH_member_array = ""
    '
	' FOR i = 0 to total_clients
	' 	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
	' 		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
	' 			'msgbox all_clients_
	' 			HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
	' 		END IF
	' 	END IF
	' NEXT

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
	HH_member_array = SPLIT(HH_member_array, " ")
end function

function read_JOBS_panel()
'--- This function adds STAT/JOBS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
'~~~~~ JOBS_variable: the variable used by the editbox you wish to autofill.
'===== Keywords: MAXIS, autofill, JOBS
    EMReadScreen JOBS_month, 5, 20, 55									'reads Footer month
    JOBS_month = replace(JOBS_month, " ", "/")					'Cleans up the read number by putting a / in place of the blank space between MM YY
    EMReadScreen JOBS_type, 30, 7, 42										'Reads up name of the employer and then cleans it up
    JOBS_type = replace(JOBS_type, "_", ""	)
    JOBS_type = trim(JOBS_type)
    JOBS_type = split(JOBS_type)
    For each JOBS_part in JOBS_type											'Correcting case on the name of the employer as it reads in all CAPS
    If JOBS_part <> "" then
        first_letter = ucase(left(JOBS_part, 1))
        other_letters = LCase(right(JOBS_part, len(JOBS_part) -1))
        new_JOBS_type = new_JOBS_type & first_letter & other_letters & " "
    End if
    Next
    ALL_JOBS_PANELS_ARRAY(employer_name, job_count) = new_JOBS_type
    EMReadScreen jobs_hourly_wage, 6, 6, 75   'reading hourly wage field
    ALL_JOBS_PANELS_ARRAY(hrly_wage, job_count) = replace(jobs_hourly_wage, "_", "")   'trimming any underscores

    ' Navigates to the FS PIC
    EMWriteScreen "x", 19, 38
    transmit
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, job_count) = trim(SNAP_JOBS_amt)
    EMReadScreen jobs_SNAP_prospective_amt, 8, 18, 56
    ALL_JOBS_PANELS_ARRAY(pic_prosp_income, job_count) = trim(jobs_SNAP_prospective_amt)  'prospective amount from PIC screen
    EMReadScreen snap_pay_frequency, 1, 5, 64
    ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = snap_pay_frequency
    EMReadScreen date_of_pic_calc, 8, 5, 34
    ALL_JOBS_PANELS_ARRAY(pic_calc_date, job_count) = replace(date_of_pic_calc, " ", "/")
    transmit
    'Navigats to GRH PIC
    EMReadscreen GRH_PIC_check, 3, 19, 73 	'This must check to see if the GRH PIC is there or not. If fun on months 06/16 and before it will cause an error if it pf3s on the home panel.
    IF GRH_PIC_check = "GRH" THEN
    	EMWriteScreen "x", 19, 71
    	transmit
    	EMReadScreen GRH_JOBS_amt, 8, 16, 69
    	GRH_JOBS_amt = trim(GRH_JOBS_amt)
    	EMReadScreen GRH_pay_frequency, 1, 3, 63
    	EMReadScreen GRH_date_of_pic_calc, 8, 3, 30
    	GRH_date_of_pic_calc = replace(GRH_date_of_pic_calc, " ", "/")
    	PF3
    END IF
    '  Reads the information on the retro side of JOBS
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    ALL_JOBS_PANELS_ARRAY(job_retro_income, job_count) = trim(retro_JOBS_amt)

    '  Reads the information on the prospective side of JOBS
    EMReadScreen prospective_JOBS_amt, 8, 17, 67
    ALL_JOBS_PANELS_ARRAY(job_prosp_income, job_count) = trim(prospective_JOBS_amt)
    '  Reads the information about health care off of HC Income Estimator
    EMReadScreen pay_frequency, 1, 18, 35
    EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
    IF HC_income_est_check = "Est" Then 'this is the old position
    	EMWriteScreen "x", 19, 54
    ELSE								'this is the new position
    	EMWriteScreen "x", 19, 48
    END IF
    transmit
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    transmit

    EMReadScreen JOBS_ver, 25, 6, 36
    ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = trim(JOBS_ver)
    EMReadScreen JOBS_income_end_date, 8, 9, 49
    'This now cleans up the variables converting codes read from the panel into words for the final variable to be used in the output.
    If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
    If IsDate(JOBS_income_end_date) = True then
        variable_name_for_JOBS = variable_name_for_JOBS & new_JOBS_type & "(ended " & JOBS_income_end_date & "); "
    Else
        If pay_frequency = "1" then pay_frequency = "monthly"
        If pay_frequency = "2" then pay_frequency = "semimonthly"
        If pay_frequency = "3" then pay_frequency = "biweekly"
        If pay_frequency = "4" then pay_frequency = "weekly"
        If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
        IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
        IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
        IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
        IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
        IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
        If GRH_pay_frequency = "1" then GRH_pay_frequency = "monthly"
        If GRH_pay_frequency = "2" then GRH_pay_frequency = "semimonthly"
        If GRH_pay_frequency = "3" then GRH_pay_frequency = "biweekly"
        If GRH_pay_frequency = "4" then GRH_pay_frequency = "weekly"
        variable_name_for_JOBS = variable_name_for_JOBS & "EI from " & trim(new_JOBS_type) & ", " & JOBS_month  & " amts:; "
        If SNAP_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- SNAP PIC: $" & SNAP_JOBS_amt & "/" & snap_pay_frequency & ", SNAP PIC Prospective: $" & jobs_SNAP_prospective_amt & ", calculated " & date_of_pic_calc & "; "
        If GRH_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- GRH PIC: $" & GRH_JOBS_amt & "/" & GRH_pay_frequency & ", calculated " & GRH_date_of_pic_calc & "; "
        If retro_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- Retrospective: $" & retro_JOBS_amt & " total; "
        IF prospective_JOBS_amt <> "" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Prospective: $" & prospective_JOBS_amt & " total; "
        IF isnumeric(jobs_hourly_wage) THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Hourly Wage: $" & jobs_hourly_wage & "; "
        'Leaving out HC income estimator if footer month is not Current month + 1
        current_month_for_hc_est = dateadd("m", "1", date)
        current_month_for_hc_est = datepart("m", current_month_for_hc_est)
        IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
        IF MAXIS_footer_month = current_month_for_hc_est THEN
            IF HC_JOBS_amt <> "________" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- HC Inc Est: $" & HC_JOBS_amt & "/" & pay_frequency & "; "
        END IF
        If JOBS_ver = "N" or JOBS_ver = "?" then variable_name_for_JOBS = variable_name_for_JOBS & "- No proof provided for this panel; "
    End if
end function

'===========================================================================================================================

'JOBS ARRAY CONSTANTS
const memb_numb     = 0
const job_instance  = 1
const employer_name = 2
Const estimate_only = 3
const verif_explain = 4
const verif_code    = 5
const info_month    = 6
const hrly_wage     = 7
const job_retro_income      = 8
const job_prosp_income      = 9
const pic_pay_date_income   = 10
const pic_pay_freq          = 11
const pic_prosp_income      = 12
const pic_calc_date         = 13
const EI_case_note          = 14
const budget_explain        = 15

Dim ALL_JOBS_PANELS_ARRAY()
ReDim ALL_JOBS_PANELS_ARRAY(budget_explain, 0)

const clt_name      = 1
const clt_age       = 2
const include_cash_checkbox     = 3
const include_snap_checkbox     = 4
const include_emer_checkbox     = 5
const count_cash_checkbox       = 6
const count_snap_checkbox       = 7
const count_emer_checkbox       = 8
const clt_notes                 = 9

Dim ALL_MEMBERS_ARRAY()
ReDim ALL_MEMBERS_ARRAY(clt_notes, 0)
'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_checkbox = checked 'The script should default to having the application signed.

'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
get_county_code				'since there is a county specific checkbox, this makes the the county clear
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

BeginDialog Dialog1, 0, 0, 281, 215, "Case number dialog"
  EditBox 65, 50, 60, 15, MAXIS_case_number
  EditBox 210, 50, 15, 15, MAXIS_footer_month
  EditBox 230, 50, 15, 15, MAXIS_footer_year
  CheckBox 10, 85, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 50, 85, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 90, 85, 35, 10, "EMER", EMER_on_CAF_checkbox
  DropListBox 185, 80, 75, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Recertification"+chr(9)+"Addendum", CAF_type
  EditBox 40, 130, 220, 15, cash_other_req_detail
  EditBox 40, 150, 220, 15, snap_other_req_detail
  EditBox 40, 170, 220, 15, emer_other_req_detail
  ButtonGroup ButtonPressed
    OkButton 170, 195, 50, 15
    CancelButton 225, 195, 50, 15
    PushButton 10, 30, 105, 10, "NOTES - Interview Completed", interview_completed_button
  Text 10, 10, 265, 20, "This script works best when run AFTER all STAT panels have been updated. If STAT panels have not been updated but you need to case note the interview use "
  Text 10, 55, 50, 10, "Case number:"
  Text 140, 55, 65, 10, "Footer month/year: "
  GroupBox 5, 70, 125, 30, "Programs marked on CAF"
  Text 145, 85, 35, 10, "CAF type:"
  GroupBox 5, 105, 265, 85, "OTHER Program Requests (not marked on CAF)"
  Text 40, 120, 130, 10, "Explain how the program was reuested."
  Text 15, 135, 20, 10, "Cash:"
  Text 15, 155, 20, 10, "SNAP:"
  Text 15, 175, 25, 10, "EMER:"
EndDialog

'initial dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		If buttonpressed = interview_completed_button Then
            confirm_run_another_script = MsgBox("You have selected the 'NOTES - Interview completed' option. This will stop the NOTES - CAF script and run the script NOTES - Interview Completed." & vbNewLine & vbNewLine &_
                                                "This option is best for when the STAT panels have not been updated when running the script. We recommend runing NOTES - CAF once STAT panels are updated to capture the correct case information in CASE/NOTE." & vbNewLine & vbNewLine &_
                                                "Would you like to continue to NOTES - Interview Completed?", vbQuestion + vbYesNo, "Stop CAF Script?")
            If confirm_run_another_script = vbYes Then Call run_from_GitHub(script_repository & "notes/interview-completed.vbs")
            If confirm_run_another_script = vbNo Then err_msg = "LOOP" & err_msg
        End If

        If CAF_type = "Select One:" then err_msg = err_msg & vbnewline & "* You must select the CAF type."
        Call validate_MAXIS_case_number(err_msg, "*")
        If IsNumeric(MAXIS_footer_month) = FALSE OR len(MAXIS_footer_month) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid Footer Month."
        If IsNumeric(MAXIS_footer_year) = FALSE OR len(MAXIS_footer_year) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid Footer Year."
        If CASH_on_CAF_checkbox = unchecked AND SNAP_on_CAF_checkbox = unchecked AND EMER_on_CAF_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* At least one program should be marked on the CAF."
        If CASH_on_CAF_checkbox = checked AND trim(cash_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."
        If SNAP_on_CAF_checkbox = checked AND trim(snap_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."
        If EMER_on_CAF_checkbox = checked AND trim(emer_other_req_detail) <> "" Then err_msg = err_msg & vbNewLine & "* If CASH was marked on the CAF, then another way of requesting does not need to be indicated."

        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If CASH_on_CAF_checkbox = checked or trim(cash_other_req_detail) <> "" Then cash_checkbox = checked
If SNAP_on_CAF_checkbox = checked or trim(snap_other_req_detail) <> "" Then SNAP_checkbox = checked
If EMER_on_CAF_checkbox = checked or trim(emer_other_req_detail) <> "" Then EMER_checkbox = checked

MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
MAXIS_footer_year = right("00" & MAXIS_footer_year, 2)
call check_for_MAXIS(False)	'checking for an active MAXIS session
MAXIS_footer_month_confirmation	'function will check the MAXIS panel footer month/year vs. the footer month/year in the dialog, and will navigate to the dialog month/year if they do not match.

'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact an alpha user for your agency.")

'Creating a custom dialog for determining who the HH members are
call HH_comp_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
	call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
	IF SNAP_checkbox = checked THEN																															'checking for SNAP 24 month renewals.'
		EMWriteScreen "X", 05, 58																																	'opening the FS revw screen.
		transmit
		EMReadScreen SNAP_recert_date, 8, 9, 64
		PF3
		SNAP_recert_date = replace(SNAP_recert_date, " ", "/")																		'replacing the read blank spaces with / to make it a date
		SNAP_recert_compare_date = dateadd("m", "12", MAXIS_footer_month & "/01/" & MAXIS_footer_year)		'making a dummy variable to compare with, by adding 12 months to the requested footer month/year.
		IF datediff("d", SNAP_recert_compare_date, SNAP_recert_date) > 0 THEN											'If the read recert date is more than 0 days away from 12 months plus the MAXIS footer month/year then it is likely a 24 month period.'
			SNAP_recert_is_likely_24_months = TRUE
		ELSE
			SNAP_recert_is_likely_24_months = FALSE																									'otherwise if we don't we set it as false
		END IF
	END IF
Else
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
End if
If HC_checkbox = checked and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")

'I put these sections in here, just because SHEL should come before HEST, it just looks cleaner.
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)

'Now it grabs the rest of the info, not dependent on which programs are selected.
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "FMED", FMED)
call autofill_editbox_from_MAXIS(HH_member_array, "IMIG", IMIG)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
'call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)

job_count = 0
call navigate_to_MAXIS_screen("STAT", "JOBS")
EMReadScreen panel_total_check, 6, 2, 73
IF panel_total_check <> "0 Of 0" Then
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen JOBS_total, 1, 2, 78
        If JOBS_total <> 0 then
          variable_written_to = variable_written_to & "Member " & HH_member & "- "
          Do
            ReDim Preserve ALL_JOBS_PANELS_ARRAY(budget_explain, job_count)
            ALL_JOBS_PANELS_ARRAY(memb_numb, job_count) = HH_member
            ALL_JOBS_PANELS_ARRAY(info_month, job_count) = MAXIS_footer_month & "/" & MAXIS_footer_year
            call read_JOBS_panel

            EMReadScreen JOBS_panel_current, 1, 2, 73
            ALL_JOBS_PANELS_ARRAY(job_instance, job_count) = "0" & JOBS_panel_current

            If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
            job_count = job_count + 1
          Loop until cint(JOBS_panel_current) = cint(JOBS_total)
        End if
    Next
End If
'FOR EACH JOB PANEL GO LOOK FOR A RECENT EI CASE NOTE'
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "CASH, "
If HC_checkbox = checked then programs_applied_for = programs_applied_for & "HC, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

'SHOULD DEFAULT TO TIKLING FOR APPLICATIONS THAT AREN'T RECERTS.
If CAF_type <> "Recertification" then TIKL_checkbox = checked

'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Do
	Do
		Do
			Do

                BeginDialog Dialog1, 0, 0, 451, 290, "CAF dialog part 1"
                  EditBox 60, 5, 50, 15, CAF_datestamp
                  ComboBox 175, 5, 70, 15, " "+chr(9)+"phone"+chr(9)+"office", interview_type
                  CheckBox 255, 5, 65, 10, "Used Interpreter", Used_Interpreter_checkbox
                  EditBox 60, 25, 50, 15, interview_date
                  ComboBox 230, 25, 95, 15,  " "+chr(9)+"Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Office"+chr(9)+"Online", how_app_rcvd
                  ComboBox 220, 45, 105, 15, " "+chr(9)+"DHS-2128 (LTC Renewal)"+chr(9)+"DHS-3417B (Req. to Apply...)"+chr(9)+"DHS-3418 (HC Renewal)"+chr(9)+"DHS-3531 (LTC Application)"+chr(9)+"DHS-3876 (Certain Pops App)"+chr(9)+"DHS-6696(MNsure HC App)", HC_document_received
                  EditBox 390, 45, 50, 15, HC_datestamp
                  EditBox 75, 70, 370, 15, HH_comp
                  EditBox 35, 90, 200, 15, cit_id
                  EditBox 265, 90, 180, 15, IMIG
                  EditBox 60, 110, 120, 15, AREP
                  EditBox 270, 110, 175, 15, SCHL
                  EditBox 60, 130, 210, 15, DISA
                  EditBox 310, 130, 135, 15, FACI
                  EditBox 35, 160, 410, 15, PREG
                  EditBox 35, 180, 410, 15, ABPS
                  EditBox 35, 200, 410, 15, EMPS
                  If worker_county_code = "x127" or worker_county_code = "x162" then CheckBox 35, 220, 180, 10, "Sent MFIP financial orientation DVD to participant(s).", MFIP_DVD_checkbox
                  EditBox 55, 235, 390, 15, verifs_needed
                  ButtonGroup ButtonPressed
                    PushButton 340, 270, 50, 15, "NEXT", next_to_page_02_button
                    CancelButton 395, 270, 50, 15
                    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                    PushButton 335, 25, 45, 10, "next panel", next_panel_button
                    PushButton 395, 25, 45, 10, "next memb", next_memb_button
                    PushButton 5, 75, 60, 10, "HH comp/EATS:", EATS_button
                    PushButton 240, 95, 20, 10, "IMIG:", IMIG_button
                    PushButton 5, 115, 25, 10, "AREP/", AREP_button
                    PushButton 30, 115, 25, 10, "ALTP:", ALTP_button
                    PushButton 190, 115, 25, 10, "SCHL/", SCHL_button
                    PushButton 215, 115, 25, 10, "STIN/", STIN_button
                    PushButton 240, 115, 25, 10, "STEC:", STEC_button
                    PushButton 5, 135, 25, 10, "DISA/", DISA_button
                    PushButton 30, 135, 25, 10, "PDED:", PDED_button
                    PushButton 280, 135, 25, 10, "FACI:", FACI_button
                    PushButton 5, 165, 25, 10, "PREG:", PREG_button
                    PushButton 5, 185, 25, 10, "ABPS:", ABPS_button
                    PushButton 5, 205, 25, 10, "EMPS", EMPS_button
                    PushButton 10, 270, 20, 10, "DWP", ELIG_DWP_button
                    PushButton 30, 270, 15, 10, "FS", ELIG_FS_button
                    PushButton 45, 270, 15, 10, "GA", ELIG_GA_button
                    PushButton 60, 270, 15, 10, "HC", ELIG_HC_button
                    PushButton 75, 270, 20, 10, "MFIP", ELIG_MFIP_button
                    PushButton 95, 270, 20, 10, "MSA", ELIG_MSA_button
                    PushButton 130, 270, 25, 10, "ADDR", ADDR_button
                    PushButton 155, 270, 25, 10, "MEMB", MEMB_button
                    PushButton 180, 270, 25, 10, "MEMI", MEMI_button
                    PushButton 205, 270, 25, 10, "PROG", PROG_button
                    PushButton 230, 270, 25, 10, "REVW", REVW_button
                    PushButton 255, 270, 25, 10, "SANC", SANC_button
                    PushButton 280, 270, 25, 10, "TIME", TIME_button
                    PushButton 305, 270, 25, 10, "TYPE", TYPE_button
                  Text 335, 50, 55, 10, "HC datestamp:"
                  Text 5, 95, 25, 10, "CIT/ID:"
                  Text 5, 240, 50, 10, "Verifs needed:"
                  GroupBox 5, 260, 115, 25, "ELIG panels:"
                  GroupBox 125, 260, 210, 25, "other STAT panels:"
                  GroupBox 330, 5, 115, 35, "STAT-based navigation"
                  Text 5, 10, 55, 10, "CAF datestamp:"
                  Text 5, 30, 55, 10, "Interview date:"
                  Text 120, 10, 50, 10, "Interview type:"
                  Text 5, 50, 210, 10, "If HC applied for (or recertifying): what document was received?:"
                  Text 120, 30, 110, 10, "How was application received?:"
                EndDialog

				err_msg = ""
				Dialog Dialog1			'Displays the first dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				If CAF_datestamp = "" or len(CAF_datestamp) > 10 THEN err_msg = "Please enter a valid application datestamp."
				If err_msg <> "" THEN Msgbox err_msg
			Loop until ButtonPressed = next_to_page_02_button and err_msg = ""

            ' 'THIS IS THE BASE FOR AN INCOME SPECIFIC DIALOG'
            ' BeginDialog Dialog1, 0, 0, 606, 145, "CAF Income Information"
            '   GroupBox 5, 5, 595, 115, "Earned Income"
            '   Text 15, 20, 170, 10, "Member 01 - EMPLOYER"
            '   CheckBox 375, 20, 220, 10, "Check here if this income is not verified and is only an estimate.", no_verif_estimate_only_checkbox
            '   Text 35, 40, 40, 10, "Verification:"
            '   EditBox 85, 35, 240, 15, verif_explanation
            '   Text 340, 40, 75, 10, "Footer Month: XX/XX"
            '   Text 425, 40, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
            '   Text 35, 60, 45, 10, "Hourly Wage:"
            '   EditBox 85, 55, 50, 15, hourly_wage
            '   Text 150, 60, 75, 10, "Retrospective Income:"
            '   EditBox 230, 55, 125, 15, retro_income
            '   Text 360, 60, 70, 10, "Prospective Income:"
            '   EditBox 430, 55, 165, 15, prosp_income
            '   Text 35, 80, 35, 10, "SNAP PIC:"
            '   Text 75, 80, 60, 10, "Pay Date Amount: "
            '   EditBox 135, 75, 50, 15, pay_date_amount
            '   ComboBox 195, 75, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", pay_date_frequency
            '   Text 285, 80, 70, 10, "Prospective Amount:"
            '   EditBox 360, 75, 60, 15, pic_prosp_income
            '   Text 440, 80, 40, 10, "Calculated:"
            '   EditBox 490, 75, 50, 15, calculated_date
            '   Text 35, 100, 55, 10, "Explain Budget:"
            '   EditBox 90, 95, 505, 15, budget_explanation
            '   ButtonGroup ButtonPressed
            '     PushButton 430, 130, 60, 10, "previous page", previous_to_page_01_button
            '     PushButton 495, 125, 50, 15, "NEXT", next_to_page_03_button
            '     CancelButton 550, 125, 50, 15
            '     PushButton 15, 130, 25, 10, "BUSI", BUSI_button
            '     PushButton 45, 130, 25, 10, "JOBS", JOBS_button
            '     PushButton 75, 130, 25, 10, "PBEN", PBEN_button
            '     PushButton 105, 130, 25, 10, "RBIC", RBIC_button
            '     PushButton 135, 130, 25, 10, "UNEA", UNEA_button
            '     PushButton 175, 130, 45, 10, "prev. panel", prev_panel_button
            '     PushButton 225, 130, 45, 10, "next panel", next_panel_button
            '     PushButton 275, 130, 45, 10, "prev. memb", prev_memb_button
            '     PushButton 325, 130, 45, 10, "next memb", next_memb_button
            ' EndDialog
            '
            ' BeginDialog Dialog1, 0, 0, 606, 245, "CAF Income Information"
            '   GroupBox 5, 5, 595, 215, "Earned Income"
            '   Text 15, 20, 170, 10, "Member 01 - EMPLOYER"
            '   CheckBox 375, 20, 220, 10, "Check here if this income is not verified and is only an estimate.", no_verif_estimate_only_checkbox
            '   Text 35, 40, 40, 10, "Verification:"
            '   EditBox 85, 35, 240, 15, verif_explanation
            '   Text 340, 40, 75, 10, "Footer Month: XX/XX"
            '   Text 425, 40, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
            '   Text 35, 60, 45, 10, "Hourly Wage:"
            '   EditBox 85, 55, 50, 15, hourly_wage
            '   Text 150, 60, 75, 10, "Retrospective Income:"
            '   EditBox 230, 55, 125, 15, retro_income
            '   Text 360, 60, 70, 10, "Prospective Income:"
            '   EditBox 430, 55, 165, 15, prosp_income
            '   Text 35, 80, 35, 10, "SNAP PIC:"
            '   Text 75, 80, 60, 10, "Pay Date Amount: "
            '   EditBox 135, 75, 50, 15, pay_date_amount
            '   ComboBox 195, 75, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", pay_date_frequency
            '   Text 285, 80, 70, 10, "Prospective Amount:"
            '   EditBox 360, 75, 60, 15, pic_prosp_income
            '   Text 440, 80, 40, 10, "Calculated:"
            '   EditBox 490, 75, 50, 15, calculated_date
            '   Text 35, 100, 55, 10, "Explain Budget:"
            '   EditBox 90, 95, 505, 15, budget_explanation
            '   ButtonGroup ButtonPressed
            '     PushButton 430, 230, 60, 10, "previous page", previous_to_page_01_button
            '     PushButton 495, 225, 50, 15, "NEXT", next_to_page_03_button
            '     CancelButton 550, 225, 50, 15
            '     PushButton 15, 230, 25, 10, "BUSI", BUSI_button
            '     PushButton 45, 230, 25, 10, "JOBS", JOBS_button
            '     PushButton 75, 230, 25, 10, "PBEN", PBEN_button
            '     PushButton 105, 230, 25, 10, "RBIC", RBIC_button
            '     PushButton 135, 230, 25, 10, "UNEA", UNEA_button
            '     PushButton 175, 230, 45, 10, "prev. panel", prev_panel_button
            '     PushButton 225, 230, 45, 10, "next panel", next_panel_button
            '     PushButton 275, 230, 45, 10, "prev. memb", prev_memb_button
            '     PushButton 325, 230, 45, 10, "next memb", next_memb_button
            '   Text 15, 125, 170, 10, "Member 01 - EMPLOYER"
            '   CheckBox 375, 125, 220, 10, "Check here if this income is not verified and is only an estimate.", Check2
            '   Text 35, 145, 40, 10, "Verification:"
            '   EditBox 85, 140, 240, 15, Edit9
            '   Text 340, 145, 75, 10, "Footer Month: XX/XX"
            '   Text 425, 145, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
            '   Text 35, 165, 45, 10, "Hourly Wage:"
            '   EditBox 85, 160, 50, 15, Edit10
            '   Text 150, 165, 75, 10, "Retrospective Income:"
            '   EditBox 230, 160, 125, 15, Edit11
            '   Text 360, 165, 70, 10, "Prospective Income:"
            '   EditBox 430, 160, 165, 15, Edit12
            '   Text 35, 185, 35, 10, "SNAP PIC:"
            '   Text 75, 185, 60, 10, "Pay Date Amount: "
            '   EditBox 135, 180, 50, 15, Edit13
            '   ComboBox 195, 180, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", Combo2
            '   Text 285, 185, 70, 10, "Prospective Amount:"
            '   EditBox 360, 180, 60, 15, Edit14
            '   Text 440, 185, 40, 10, "Calculated:"
            '   EditBox 490, 180, 50, 15, Edit15
            '   Text 35, 205, 55, 10, "Explain Budget:"
            '   EditBox 90, 200, 505, 15, Edit16
            ' EndDialog


            Do
                err_msg = ""

                dlg_len = 45
                jobs_grp_len = 15
                'NEED HANDLING FOR IF NO JOBS'
                For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                    dlg_len = dlg_len + 100
                    jobs_grp_len = jobs_grp_len + 100
                Next
                y_pos = 20

                BeginDialog Dialog1, 0, 0, 606, dlg_len, "CAF Income Information"
                  GroupBox 5, 5, 595, jobs_grp_len, "Earned Income"
                  For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                      Text 15, y_pos, 150, 10, "Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
                      Text 170, y_pos, 200, 10, "Verif: " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job)
                      CheckBox 375, y_pos, 220, 10, "Check here if this income is not verified and is only an estimate.", ALL_JOBS_PANELS_ARRAY(estimate_only, each_job)
                      y_pos = y_pos + 20
                      Text 35, y_pos, 40, 10, "Verification:"
                      EditBox 85, y_pos - 5, 240, 15, ALL_JOBS_PANELS_ARRAY(verif_explain, each_job)
                      Text 340, y_pos, 75, 10, "Footer Month: " & ALL_JOBS_PANELS_ARRAY(info_month, each_job)
                      IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then Text 425, y_pos, 175, 10, "EARNED INCOME BUDGETING CASE NOTE FOUND"
                      y_pos = y_pos + 20
                      Text 35, y_pos, 45, 10, "Hourly Wage:"
                      EditBox 85, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(hrly_wage, each_job)
                      Text 150, y_pos, 75, 10, "Retrospective Income:"
                      EditBox 230, y_pos - 5, 125, 15, ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job)
                      Text 360, y_pos, 70, 10, "Prospective Income:"
                      EditBox 430, y_pos - 5, 165, 15, ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job)
                      y_pos = y_pos + 20
                      Text 35, y_pos, 35, 10, "SNAP PIC:"
                      Text 75, y_pos, 60, 10, "Pay Date Amount: "
                      EditBox 135, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)
                      ComboBox 195, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job)
                      Text 285, y_pos, 70, 10, "Prospective Amount:"
                      EditBox 360, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)
                      Text 440, y_pos, 40, 10, "Calculated:"
                      EditBox 490, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_calc_date, each_job)
                      y_pos = y_pos + 20
                      Text 35, y_pos, 55, 10, "Explain Budget:"
                      EditBox 90, y_pos - 5, 505, 15, ALL_JOBS_PANELS_ARRAY(budget_explain, each_job)
                      y_pos = y_pos + 25
                  Next
                  ButtonGroup ButtonPressed
                    PushButton 430, y_pos, 60, 10, "previous page", previous_to_page_01_button
                    PushButton 495, y_pos - 5, 50, 15, "NEXT", next_to_page_03_button
                    CancelButton 550, y_pos - 5, 50, 15
                    PushButton 15, y_pos, 25, 10, "BUSI", BUSI_button
                    PushButton 45, y_pos, 25, 10, "JOBS", JOBS_button
                    PushButton 75, y_pos, 25, 10, "PBEN", PBEN_button
                    PushButton 105, y_pos, 25, 10, "RBIC", RBIC_button
                    PushButton 135, y_pos, 25, 10, "UNEA", UNEA_button
                    PushButton 175, y_pos, 45, 10, "prev. panel", prev_panel_button
                    PushButton 225, y_pos, 45, 10, "next panel", next_panel_button
                    PushButton 275, y_pos, 45, 10, "prev. memb", prev_memb_button
                    PushButton 325, y_pos, 45, 10, "next memb", next_memb_button
                EndDialog

                dialog Dialog1
                cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)

                For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                    'err_msg'
                    'IF THERE IS AN EI CASE NOTE - DON'T WORRY ABOUT MUCH ERR HANDLING
                Next


            Loop until ButtonPressed = next_to_page_03_button and err_msg = ""

			Do
				Do      'THIS NEEDS TO BE WORKED ON'
                    BeginDialog Dialog1, 0, 0, 451, 305, "CAF dialog part 2"
                      EditBox 65, 35, 375, 15, notes_on_abawd
                      EditBox 65, 95, 375, 15, SHEL_HEST
                      EditBox 65, 135, 375, 15, COEX_DCEX
                      EditBox 65, 175, 375, 15, CASH_ACCTs
                      EditBox 105, 235, 335, 15, other_assets
                      EditBox 55, 265, 385, 15, verifs_needed
                      ButtonGroup ButtonPressed
                        PushButton 270, 290, 60, 10, "previous page", previous_to_page_02_button
                        PushButton 335, 285, 50, 15, "NEXT", next_to_page_04_button
                        CancelButton 390, 285, 50, 15
                        PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
                        PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
                        PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
                        PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
                        PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
                        PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
                        PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
                        PushButton 240, 15, 45, 10, "prev. panel", prev_panel_button
                        PushButton 345, 15, 45, 10, "next panel", next_panel_button
                        PushButton 290, 15, 45, 10, "prev. memb", prev_memb_button
                        PushButton 395, 15, 45, 10, "next memb", next_memb_button
                        PushButton 20, 40, 35, 10, "WREG:", WREG_button
                        PushButton 25, 80, 25, 10, "SHEL:", SHEL_button
                        PushButton 25, 100, 25, 10, "HEST:", HEST_button
                        PushButton 25, 120, 25, 10, "COEX:", COEX_button
                        PushButton 25, 140, 25, 10, "DCEX:", DCEX_button
                        PushButton 340, 160, 25, 10, "CASH:", CASH_button
                        PushButton 25, 180, 30, 10, "ACCTs:", ACCT_button
                        PushButton 30, 200, 25, 10, "CARS:", CARS_button
                        PushButton 30, 220, 25, 10, "REST:", REST_button
                        PushButton 5, 240, 25, 10, "SECU/", SECU_button
                        PushButton 30, 240, 25, 10, "TRAN/", TRAN_button
                        PushButton 55, 240, 45, 10, "other assets:", OTHR_button
                      GroupBox 235, 5, 205, 25, "STAT-based navigation"
                      Text 5, 270, 50, 10, "Verifs needed:"
                      GroupBox 5, 5, 130, 25, "ELIG panels:"
                      ButtonGroup ButtonPressed
                        PushButton 20, 60, 35, 10, "ABAWD:", ABAWD_button
                      EditBox 65, 55, 375, 15, Edit12
                      EditBox 65, 75, 375, 15, Edit13
                      Text 5, 160, 60, 10, "Other Deductions:"
                      EditBox 65, 155, 260, 15, Edit14
                      EditBox 65, 115, 375, 15, Edit15
                      EditBox 370, 155, 70, 15, Edit16
                      EditBox 65, 195, 375, 15, Edit17
                      EditBox 65, 215, 375, 15, Edit18
                    EndDialog


					err_msg = ""
					income_note_error_msg = ""
					Dialog Dialog1			'Displays the second dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
					If ButtonPressed = income_notes_button Then
                        BeginDialog Dialog1, 0, 0, 351, 215, "Explanation of Income"
                          CheckBox 10, 30, 325, 10, "JOBS - Income detail on previous note(s)", see_other_note_checkbox
                          CheckBox 10, 45, 325, 10, "JOBS - Income has not been verified and detail will be entered when received.", not_verified_checkbox
                          CheckBox 10, 60, 325, 10, "JOBS - Client has confirmed that JOBS income is expected to continue at this rate and hours.", jobs_anticipated_checkbox
                          CheckBox 10, 75, 330, 10, "JOBS - This is a new job and actual check stubs are not available, advised client that if actual pay", new_jobs_checkbox
                          CheckBox 10, 100, 325, 10, "BUSI - Client has confirmed that BUSI income is expected to continue at this rate and hours.", busi_anticipated_checkbox
                          CheckBox 10, 115, 250, 10, "BUSI - Client has agreed to the self-employment budgeting method used.", busi_method_agree_checkbox
                          CheckBox 10, 130, 325, 10, "RBIC - Client has confirmed that RBIC income is expected to continue at this rate and hours.", rbic_anticipated_checkbox
                          CheckBox 10, 145, 325, 10, "UNEA - Client has confirmed that UNEA income is expected to continue at this rate and hours.", unea_anticipated_checkbox
                          CheckBox 10, 160, 315, 10, "UNEA - Client has applied for unemployment benefits but no determination made at this time.", ui_pending_checkbox
                          CheckBox 45, 170, 225, 10, "Check here to have the script set a TIKL to check UI in two weeks.", tikl_for_ui
                          CheckBox 10, 185, 150, 10, "NONE - This case has no income reported.", no_income_checkbox
                          ButtonGroup ButtonPressed
                            PushButton 240, 195, 50, 15, "Insert", add_to_notes_button
                            CancelButton 295, 195, 50, 15
                          Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
                          Text 45, 85, 315, 10, "varies significantly, client should provide proof of this difference to have benefits adjusted."
                        EndDialog

						Dialog Dialog1
						If ButtonPressed = add_to_notes_button Then
                            If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
                            If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
							If jobs_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects all income from jobs to continue at this amount."
							If new_jobs_checkbox = checked Then notes_on_income = notes_on_income & "; This is a new job and actual check stubs have not been received, advised client to provide proof once pay is received if the income received differs significantly."
							If busi_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects all income from self employment to continue at this amount."
							If busi_method_agree_checkbox = checked Then notes_on_income = notes_on_income & "; Explained to client the self employment budgeting methods and client agreed to the method used."
							If rbic_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects roomer/boarder income to continue at this amount."
							If unea_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects unearned income to continue at this amount."
							If ui_pending_checkbox = checked Then notes_on_income = notes_on_income & "; Client has applied for Unemployment Income recently but request is still pending, will need to be reviewed soon for changes."
							If tikl_for_ui = checked Then notes_on_income = notes_on_income & " TIKL set to request an update on Unemployment Income."
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; Client has reported they have no income and do not expect any changes to this at this time."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
					IF (earned_income <> "" AND trim(notes_on_income) = "") OR (unearned_income <> "" AND notes_on_income = "") THEN income_note_error_msg = True
					If err_msg <> "" THEN Msgbox err_msg
				Loop until ButtonPressed = (next_to_page_04_button AND err_msg = "") or (ButtonPressed = previous_to_page_02_button AND err_msg = "")		'If you press either the next or previous button, this loop ends
				If ButtonPressed = previous_to_page_01_button then exit do		'If the button was previous, it exits this do loop and is caught in the next one, which sends you back to Dialog 1 because of the "If ButtonPressed = previous_to_page_01_button then exit do" later on
				Do
                    BeginDialog Dialog1, 0, 0, 451, 405, "CAF dialog part 3"
                      EditBox 60, 45, 385, 15, INSA
                      EditBox 35, 65, 410, 15, ACCI
                      EditBox 35, 85, 175, 15, DIET
                      EditBox 245, 85, 200, 15, BILS
                      EditBox 35, 105, 285, 15, FMED
                      EditBox 390, 105, 55, 15, retro_request
                      EditBox 180, 130, 265, 15, reason_expedited_wasnt_processed
                      EditBox 100, 150, 345, 15, FIAT_reasons
                      CheckBox 15, 190, 80, 10, "Application signed?", application_signed_checkbox
                      CheckBox 15, 205, 65, 10, "Appt letter sent?", appt_letter_sent_checkbox
                      CheckBox 15, 220, 150, 10, "Client willing to participate with E and T", E_and_T_checkbox
                      CheckBox 15, 235, 70, 10, "EBT referral sent?", EBT_referral_checkbox
                      CheckBox 115, 190, 50, 10, "eDRS sent?", eDRS_sent_checkbox
                      CheckBox 115, 205, 50, 10, "Expedited?", expedited_checkbox
                      CheckBox 115, 235, 70, 10, "IAAs/OMB given?", IAA_checkbox
                      CheckBox 200, 190, 115, 10, "Informed client of recert period?", recert_period_checkbox
                      CheckBox 200, 205, 80, 10, "Intake packet given?", intake_packet_checkbox
                      CheckBox 200, 220, 105, 10, "Managed care packet sent?", managed_care_packet_checkbox
                      CheckBox 200, 235, 105, 10, "Managed care referral made?", managed_care_referral_checkbox
                      CheckBox 345, 190, 65, 10, "R/R explained?", R_R_checkbox
                      CheckBox 345, 205, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
                      CheckBox 345, 220, 65, 10, "Updated MMIS?", updated_MMIS_checkbox
                      CheckBox 345, 235, 95, 10, "Workforce referral made?", WF1_checkbox
                      EditBox 55, 260, 230, 15, other_notes
                      EditBox 55, 280, 390, 15, verifs_needed
                      EditBox 55, 300, 390, 15, actions_taken
                      ComboBox 330, 260, 115, 15, " "+chr(9)+"incomplete"+chr(9)+"approved", CAF_status
                      CheckBox 15, 335, 240, 10, "Check here to update PND2 to show client delay (pending cases only).", client_delay_checkbox
                      CheckBox 15, 350, 200, 10, "Check here to create a TIKL to deny at the 30/45 day mark.", TIKL_checkbox
                      CheckBox 15, 365, 265, 10, "Check here to send a TIKL (10 days from now) to update PND2 for Client Delay.", client_delay_TIKL_checkbox
                      EditBox 395, 345, 50, 15, worker_signature
                      ButtonGroup ButtonPressed
                        PushButton 290, 370, 45, 10, "prev. page", previous_to_page_03_button
                        OkButton 340, 365, 50, 15
                        CancelButton 395, 365, 50, 15
                        PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
                        PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
                        PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
                        PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
                        PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
                        PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
                        PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
                        PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                        PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                        PushButton 335, 25, 45, 10, "next panel", next_panel_button
                        PushButton 395, 25, 45, 10, "next memb", next_memb_button
                      GroupBox 5, 5, 130, 25, "ELIG panels:"
                      GroupBox 330, 5, 115, 35, "STAT-based navigation"
                      ButtonGroup ButtonPressed
                        PushButton 5, 50, 25, 10, "INSA/", INSA_button
                        PushButton 30, 50, 25, 10, "MEDI:", MEDI_button
                        PushButton 5, 70, 25, 10, "ACCI:", ACCI_button
                        PushButton 5, 90, 25, 10, "DIET:", DIET_button
                        PushButton 5, 110, 25, 10, "FMED:", FMED_button
                      Text 5, 135, 170, 10, "Reason expedited wasn't processed (if applicable):"
                      Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
                      GroupBox 5, 175, 440, 75, "Common elements workers should case note:"
                      Text 5, 265, 50, 10, "Other notes:"
                      Text 290, 265, 40, 10, "CAF status:"
                      Text 5, 285, 50, 10, "Verifs needed:"
                      Text 5, 305, 50, 10, "Actions taken:"
                      GroupBox 5, 320, 280, 60, "Actions the script can do:"
                      Text 330, 350, 60, 10, "Worker signature:"
                      ButtonGroup ButtonPressed
                        PushButton 325, 110, 60, 10, "Retro Req. date:", HCRE_button
                        PushButton 215, 90, 25, 10, "BILS:", BILS_button
                    EndDialog

					err_msg = ""
					Dialog Dialog1			'Displays the third dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
					If ButtonPressed = previous_to_page_02_button then exit do		'Exits this do...loop here if you press previous. The second ""loop until ButtonPressed = -1" gets caught, and it loops back to the "Do" after "Loop until ButtonPressed = next_to_page_02_button"
					If actions_taken = "" THEN err_msg = err_msg & vbCr & "Please complete actions taken section."    'creating err_msg if required items are missing
					If income_note_error_msg = True THEN err_msg = err_msg & VbCr & "Income for this case was found in MAXIS. Please complete the 'notes on income and budget' field."
					If worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a worker signature."
					If CAF_status = " " THEN err_msg = err_msg & vbCr & "Please select a CAF Status."
					If err_msg <> "" THEN Msgbox err_msg
				Loop until (ButtonPressed = -1 and err_msg = "") or (ButtonPressed = previous_to_page_03_button and err_msg = "")		'If OK or PREV, it exits the loop here, which is weird because the above also causes it to exit
			Loop until ButtonPressed = -1	'Because this is in here a second time, it triggers a return to the "Dialog CAF_dialog_02" line, where all those "DOs" start again!!!!!
			If ButtonPressed = previous_to_page_01_button then exit do 	'This exits this particular loop again for prev button on page 2, which sends you back to page 1!!
		Loop until err_msg = ""		'Loops all of that until those four sections are finished. Let's move that over to those particular pages. Folks would be less angry that way I bet.
		CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
	Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Call check_for_maxis(FALSE)  'allows for looping to check for maxis after worker has complete dialog box so as not to lose a giant CAF case note if they get timed out while writing.

'This code will update the interview date in PROG.
If CAF_type <> "Recertification" AND CAF_type <> "Addendum" Then        'Interview date is not on PROG for recertifications or addendums
    If SNAP_checkbox = checked OR cash_checkbox = checked Then          'Interviews are only required for Cash and SNAP
        intv_date_needed = FALSE
        Call navigate_to_MAXIS_screen("STAT", "PROG")                   'Going to STAT to check to see if there is already an interview indicated.

        If SNAP_checkbox = checked Then                                 'If the script is being run for a SNAP interview
            EMReadScreen entered_intv_date, 8, 10, 55                   'REading what is entered in the SNAP interview
            'MsgBox "SNAP interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE  'If this is blank - the script needs to prompt worker to update it
        End If

        If cash_checkbox = checked THen                             'If the script is bring run for a Cash interview
            EMReadScreen cash_one_app, 8, 6, 33                     'First the script needs to identify if it is cash 1 or cash 2 that has the application information
            EMReadScreen cash_two_app, 8, 7, 33
            EMReadScreen grh_cash_app, 8, 9, 33

            cash_one_app = replace(cash_one_app, " ", "/")          'Turning this in to a date format
            cash_two_app = replace(cash_two_app, " ", "/")
            grh_cash_app = replace(grh_cash_app, " ", "/")

            If cash_one_app <> "__/__/__" Then      'Error handling - VB doesn't like date comparisons with non-dates
                if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6     'If date of application on PROG matches script date of applicaton
            End If
            If cash_two_app <> "__/__/__" Then
                if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
            End If

            If grh_cash_app <> "__/__/__" Then
                if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
            End If

            EMReadScreen entered_intv_date, 8, prog_row, 55                     'Reading the right interview date with row defined above
            'MsgBox "Cash interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE      'If this is blank - script needs to prompt worker to have it updated
        End If

        If intv_date_needed = TRUE Then         'If previous code has determined that PROG needs to be updated
            If SNAP_checkbox = checked Then prog_update_SNAP_checkbox = checked     'Auto checking based on the programs the script is being run for.
            If cash_checkbox = checked Then prog_update_cash_checkbox = checked

            'Dialog code
            BeginDialog Dialog1, 0, 0, 231, 130, "Update PROG?"
              OptionGroup RadioGroup1
                RadioButton 10, 10, 155, 10, "YES! Update PROG with the Interview Date", confirm_update_prog
                RadioButton 10, 60, 90, 10, "No, do not update PROG", do_not_update_prog
              EditBox 165, 5, 50, 15, interview_date
              CheckBox 25, 25, 30, 10, "SNAP", prog_update_SNAP_checkbox
              CheckBox 25, 40, 30, 10, "CASH", prog_update_cash_checkbox
              Text 20, 75, 200, 10, "Reason PROG should not be updated with the Interview Date:"
              EditBox 20, 90, 195, 15, no_update_reason
              ButtonGroup ButtonPressed
                OkButton 175, 110, 50, 15
            EndDialog

            'Running the dialog
            Do
                err_msg = ""
                Dialog Dialog1
                'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                If do_not_update_prog = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
                IF confirm_update_prog = 1 AND prog_update_SNAP_checkbox = unchecked AND prog_update_cash_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select either CASH or SNAP to have updated on PROG."

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            If confirm_update_prog = 1 Then     'If the dialog selects to have PROG updated
                CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month

                keep_footer_month = MAXIS_footer_month      'Saving the footer month and year that was determined earlier in the script. It needs t obe changed for nav functions to work correctly
                keep_footer_year = MAXIS_footer_year

                app_month = DatePart("m", CAF_datestamp)    'Setting the footer month and year to the app month.
                app_year = DatePart("yyyy", CAF_datestamp)

                MAXIS_footer_month = right("00" & app_month, 2)
                MAXIS_footer_year = right(app_year, 2)

                CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
                PF9                                             'Edit

                intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
                intv_day = DatePart("d", interview_date)
                intv_yr = DatePart("yyyy", interview_date)

                intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
                intv_day = right("00"&intv_day, 2)
                intv_yr = right(intv_yr, 2)

                If prog_update_SNAP_checkbox = checked Then     'If it was selected to SNAP interview to be updated
                    programs_w_interview = "SNAP"               'Setting a variable for case noting

                    EMWriteScreen intv_mo, 10, 55               'SNAP is easy because there is only one area for interview - the variables go there
                    EMWriteScreen intv_day, 10, 58
                    EMWriteScreen intv_yr, 10, 61
                End If

                If prog_update_cash_checkbox = checked Then     'If it was selected to update for Cash
                    If programs_w_interview = "" Then programs_w_interview = "CASH"     'variable for the case note
                    If programs_w_interview <> "" Then programs_w_interview = "SNAP and CASH"
                    EMReadScreen cash_one_app, 8, 6, 33     'Reading app dates of both cash lines
                    EMReadScreen cash_two_app, 8, 7, 33
                    EMReadScreen grh_cash_app, 8, 9, 33

                    cash_one_app = replace(cash_one_app, " ", "/")      'Formatting as dates
                    cash_two_app = replace(cash_two_app, " ", "/")
                    grh_cash_app = replace(grh_cash_app, " ", "/")

                    If cash_one_app <> "__/__/__" Then              'Comparing them to the date of application to determine which row to use
                        if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6
                    End If
                    If cash_two_app <> "__/__/__" Then
                        if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
                    End If

                    If grh_cash_app <> "__/__/__" Then
                        if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
                    End If

                    EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
                    EMWriteScreen intv_day, prog_row, 58
                    EMWriteScreen intv_yr, prog_row, 61
                End If

                transmit                                    'Saving the panel

                MAXIS_footer_month = keep_footer_month      'resetting the footer month and year so the rest of the script uses the worker identified footer month and year.
                MAXIS_footer_year = keep_footer_year
            End If
        ENd If
    End If
End If
'MsgBox "PROG Stuff done"
'MsgBox confirm_update_prog

'Now, the client_delay_checkbox business. It'll update client delay if the box is checked and it isn't a recert.
If client_delay_checkbox = checked and CAF_type <> "Recertification" then
	call navigate_to_MAXIS_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
	for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
		EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
		If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
		EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
		If PND2_HC_status_check = "P" then
			EMWriteScreen "x", PND2_row, 3
			transmit
			person_delay_row = 7
			Do
				EMReadScreen person_delay_check, 1, person_delay_row, 39
				If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
				person_delay_row = person_delay_row + 2
			Loop until person_delay_check = " " or person_delay_row > 20
			PF3
		End if
		EMReadScreen additional_app_check, 14, PND2_row + 1, 17
		If additional_app_check <> "ADDITIONAL APP" then exit for
		PND2_row = PND2_row + 1
	next
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_checkbox = unchecked		'Probably unnecessary except that it changes the case note parameters
	End if
End if

'Going to TIKL. Now using the write TIKL function
If TIKL_checkbox = checked and CAF_type <> "Recertification" then
	If cash_checkbox = checked or EMER_checkbox = checked or SNAP_checkbox = checked then
		If DateDiff ("d", CAF_datestamp, date) > 30 Then 'Error handling to prevent script from attempting to write a TIKL in the past
			MsgBox "Cannot set TIKL as CAF Date is over 30 days old and TIKL would be in the past. You must manually track."
		Else
			call navigate_to_MAXIS_screen("dail", "writ")
			call create_MAXIS_friendly_date(CAF_datestamp, 30, 5, 18)
			If cash_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "Cash/"
			If SNAP_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "SNAP/"
			If EMER_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "EMER/"
			TIKL_msg_one = Left(TIKL_msg_one, (len(TIKL_msg_one) - 1))
			TIKL_msg_one = TIKL_msg_one & " has been pending for 30 days. Evaluate for possible denial."
			Call write_variable_in_TIKL (TIKL_msg_one)
			PF3
		End If
	End if
	If HC_checkbox = checked then
		If DateDiff ("d", CAF_datestamp, date) > 45 Then 'Error handling to prevent script from attempting to write a TIKL in the past
			MsgBox "Cannot set TIKL as CAF Date is over 45 days old and TIKL would be in the past. You must manually track."
		Else
			call navigate_to_MAXIS_screen("dail", "writ")
			call create_MAXIS_friendly_date(CAF_datestamp, 45, 5, 18)
			Call write_variable_in_TIKL ("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
			PF3
		End If
	End if
End if
If client_delay_TIKL_checkbox = checked then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	Call write_variable_in_TIKL (">>>UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE<<<")
	PF3
End if
'----Here's the new bit to TIKL to APPL the CAF for CAF_datestamp if the CL fails to complete the CASH/SNAP reinstate and then TIKL again for DateAdd("D", 30, CAF_datestamp) to evaluate for possible denial.
'----IF the DatePart("M", CAF_datestamp) = MAXIS_footer_month (DatePart("M", CAF_datestamp) is converted to footer_comparo_month for the sake of comparison) and the CAF_status <> "Approved" and CAF_type is a recertification AND cash or snap is checked, then
'---------the script generates a TIKL.
footer_comparison_month = DatePart("M", CAF_datestamp)
IF len(footer_comparison_month) <> 2 THEN footer_comparison_month = "0" & footer_comparison_month
IF CAF_type = "Recertification" AND MAXIS_footer_month = footer_comparison_month AND CAF_status <> "approved" AND (cash_checkbox = checked OR SNAP_checkbox = checked) THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	start_of_next_month = DatePart("M", DateAdd("M", 1, CAF_datestamp)) & "/01/" & DatePart("YYYY", DateAdd("M", 1, CAF_datestamp))
	denial_consider_date = DateAdd("D", 30, CAF_datestamp)
	CALL create_MAXIS_friendly_date(start_of_next_month, 0, 5, 18)
	EMWriteScreen ("IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR " & CAF_datestamp), 9, 3
	EMWriteScreen ("AND TIKL FOR " & denial_consider_date & " TO EVALUATE FOR POSSIBLE DENIAL."), 10, 3
	transmit
	PF3
END IF

IF tikl_for_ui THEN
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	two_weeks_from_now = DateAdd("d", 14, date)
	call create_MAXIS_friendly_date(two_weeks_from_now, 10, 5, 18)
	call write_variable_in_TIKL ("Review client's application for Unemployment and request an update if needed.")
	PF3
END IF
'--------------------END OF TIKL BUSINESS

'Navigates to case note, and checks to make sure we aren't in inquiry.
Call start_a_blank_CASE_NOTE

'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " recert"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***" & CAF_type & CAF_status & "***")
IF move_verifs_needed = TRUE THEN
	CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
	CALL write_variable_in_CASE_NOTE("------------------------------")
End if
CALL write_bullet_and_variable_in_CASE_NOTE("CAF datestamp", CAF_datestamp)
If Used_Interpreter_checkbox = checked then
	CALL write_variable_in_CASE_NOTE("* Interview type: " & interview_type & " w/ interpreter")
Else
	CALL write_bullet_and_variable_in_CASE_NOTE("Interview type", interview_type)
End if
CALL write_bullet_and_variable_in_CASE_NOTE("Interview date", interview_date)
'If intv_date_needed = FALSE THen CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG by worker prior to script run.")
If confirm_update_prog = 1 Then CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG for " & programs_w_interview)
If do_not_update_prog = 1 Then CALL write_bullet_and_variable_in_CASE_NOTE("PROG WAS NOT UPDATED WITH INTERVIEW DATE, because", no_update_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("HC document received", HC_document_received)
CALL write_bullet_and_variable_in_CASE_NOTE("HC datestamp", HC_datestamp)
CALL write_bullet_and_variable_in_CASE_NOTE("Programs applied for", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("How CAF was received", how_app_rcvd)
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Cit/ID", cit_id)
CALL write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
CALL write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
CALL write_bullet_and_variable_in_CASE_NOTE("SCHL/STIN/STEC", SCHL)
CALL write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
CALL write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
Call write_bullet_and_variable_in_CASE_NOTE("EMPS", EMPS)
If MFIP_DVD_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MFIP financial orientation DVD sent to participant(s).")
CALL write_bullet_and_variable_in_CASE_NOTE("ABPS", ABPS)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned inc.", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Notes on income and budget", notes_on_income)
CALL write_bullet_and_variable_in_CASE_NOTE("STWK/inc. changes", income_changes)
CALL write_bullet_and_variable_in_CASE_NOTE("ABAWD Notes/WREG", notes_on_abawd)
CALL write_bullet_and_variable_in_CASE_NOTE("Is any work temporary", is_any_work_temporary)
CALL write_bullet_and_variable_in_CASE_NOTE("SHEL/HEST", SHEL_HEST)
CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
CALL write_bullet_and_variable_in_CASE_NOTE("CASH/ACCTs", CASH_ACCTs)
CALL write_bullet_and_variable_in_CASE_NOTE("Other assets", other_assets)
CALL write_bullet_and_variable_in_CASE_NOTE("INSA", INSA)
CALL write_bullet_and_variable_in_CASE_NOTE("ACCI", ACCI)
CALL write_bullet_and_variable_in_CASE_NOTE("DIET", DIET)
CALL write_bullet_and_variable_in_CASE_NOTE("BILS", BILS)
CALL write_bullet_and_variable_in_CASE_NOTE("FMED", FMED)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro Request (IF applicable)", retro_request)
IF application_signed_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
Else
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
IF appt_letter_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Appointment letter was sent before interview.")
IF EBT_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* EBT referral made for client.")
IF eDRS_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
IF expedited_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Expedited SNAP.")
CALL write_bullet_and_variable_in_CASE_NOTE("Reason expedited wasn't processed", reason_expedited_wasnt_processed)		'This is strategically placed next to expedited checkbox entry.
IF IAA_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* IAAs/OMB given to client.")
IF intake_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received intake packet.")
IF managed_care_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received managed care packet.")
IF managed_care_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Managed care referral made.")
IF R_R_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* R/R explained to client.")
IF updated_MMIS_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated MMIS.")
IF WF1_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Workforce referral made.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF E_and_T_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client is willing to participate with E&T.")
IF recert_period_checkbox = checked THEN call write_variable_in_CASE_NOTE("* Informed client of recert period.")
IF client_delay_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
CALL write_bullet_and_variable_in_CASE_NOTE("FIAT reasons", FIAT_reasons)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF SNAP_recert_is_likely_24_months = TRUE THEN					'if we determined on stat/revw that the next SNAP recert date isn't 12 months beyond the entered footer month/year
	TIKL_for_24_month = msgbox("Your SNAP recertification date is listed as " & SNAP_recert_date & " on STAT/REVW. Do you want set a TIKL on " & dateadd("m", "-1", SNAP_recert_compare_date) & " for 12 month contact?" & vbCR & vbCR & "NOTE: Clicking yes will navigate away from CASE/NOTE saving your case note.", VBYesNo)
	IF TIKL_for_24_month = vbYes THEN 												'if the select YES then we TIKL using our custom functions.
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		CALL create_MAXIS_friendly_date(dateadd("m", "-1", SNAP_recert_compare_date), 0, 5, 18)
		CALL write_variable_in_TIKL("If SNAP is open, review to see if 12 month contact letter is needed. DAIL scrubber can send 12 Month Contact Letter if used on this TIKL.")
	END IF
END IF

end_msg = "Success! CAF has been successfully noted. Please remember to run the Approved Programs, Closed Programs, or Denied Programs scripts if  results have been APP'd."
If do_not_update_prog = 1 Then end_msg = end_msg & vbNewLine & vbNewLine & "It was selected that PROG would NOT be updated because " & no_update_reason
script_end_procedure(end_msg)
