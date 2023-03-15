'Required for statistical purposes===============================================================================
name_of_script = "ACTIONS - ADD GRH RATE 2 TO MMIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 900                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once706743
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
call changelog_update("07/23/2019", "Updated with enhanced navigation around the PPOP panel.", "Ilse Ferris, Hennepin County")
call changelog_update("03/19/2019", "Added inhibiting functionality to ensure the 'approval cty' field is filled in on the FACI panel.", "Ilse Ferris, Hennepin County")
call changelog_update("02/22/2019", "Updated coding to support agreements that include the leap year date of 02/29/2020.", "Ilse Ferris, Hennepin County")
call changelog_update("02/22/2019", "Added additional handling around the PPOP - provider selection screen. NPI numbers may have muliple facilities and addresses. Worker will need to select the faci option.", "Ilse Ferris, Hennepin County")
call changelog_update("10/12/2018", "Added handling to ensure that service agreement dates are not more than 365 days. MMIS does not support agreements over a year old.", "Ilse Ferris, Hennepin County")
call changelog_update("08/24/2018", "Added check to the script to ensure it is started in MAXIS", "Casey Love, Hennepin County")
call changelog_update("08/10/2018", "Updated to support a variety of start and end dates, includes navigation, and extra data validation to ensure cases are input correctly into MMIS.", "Ilse Ferris, Hennepin County")
call changelog_update("02/21/2018", "Added VND2 confirmation handling.", "Ilse Ferris, Hennepin County")
call changelog_update("02/12/2018", "Added out-of-county handling.", "Ilse Ferris, Hennepin County")
call changelog_update("02/08/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function HCRE_panel_bypass()
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(true)
get_county_code
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 216, 170, "Add GRH Rate 2 to MMIS"
  EditBox 110, 10, 50, 15, MAXIS_case_number
  EditBox 110, 30, 20, 15, MAXIS_footer_month
  EditBox 140, 30, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 75, 55, 40, 15
    CancelButton 120, 55, 40, 15
  Text 10, 130, 195, 25, "Before you use the script, you must have approved GRH results that reflect the SSR information in the SSR pop-up on ELIG/GRFB for the selected footer month/year."
  Text 55, 15, 50, 10, "Case Number:"
  Text 10, 95, 195, 25, "This script is to be used when a new service agreement needs to be added into MMIS. If you need to update an agreement, please do that manually."
  Text 45, 35, 60, 10, "Initial month/year:"
  GroupBox 5, 80, 205, 85, "Add GRH Rate 2 to MMIS script:"
EndDialog

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "Enter a valid case number."		'mandatory field
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
If PRIV_check = "PRIV" then script_end_procedure("PRIV case, cannot access/update. The script will now end.")

EMReadScreen grh_status, 4, 9, 74		'Ensuring that the case is active on GRH. If not, case will not be updated in MMIS.
If grh_status <> "ACTV" then
    If trim(grh_status) = "" then grh_status = "Inactive"
	script_end_procedure("GRH case status is " & grh_status & ". The script will now end.")
End if

EMReadscreen current_county, 4, 21, 21
If current_county <> UCase(worker_county_code) then script_end_procedure("Out-of-county case. Cannot update. The script will now end.")

Call HCRE_panel_bypass			'Function to bypass a jenky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.

'----------------------------------------------------------------------------------------------------UNEA panel
SSA_disa = ""
Call navigate_to_MAXIS_screen("STAT", "UNEA")
Call write_value_and_transmit("01", 20, 76)
Call write_value_and_transmit("01", 20, 79)
EMReadScreen total_panels, 1, 2, 78
If total_panels = "0" then
    SSA_disa = false
Else
    Do
        EmReadscreen UNEA_type, 2, 5, 37
        If UNEA_type = "01" or UNEA_type = "02" or UNEA_type = "03" then
            SSA_disa = True
            exit do
        ELSE
            SSA_disa = False
            transmit
        End if
        EmReadscreen error_check, 5, 24, 2
    Loop until error_check = "ENTER"
End if

'----------------------------------------------------------------------------------------------------DISA panel'
Call navigate_to_MAXIS_screen("STAT", "DISA")
Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01.
EMReadScreen waiver_type, 1, 14, 59
If waiver_type <> "_" then script_end_procedure("Client is active on a waiver. Should not be Rate 2. Please review waiver information in MMIS, and update MAXIS if applicable. The script will now end.")

If SSA_disa = True then
    EmReadscreen cert_start_date, 10, 7, 47
    EmReadscreen cert_end_date, 10, 7, 69
    If (SSA_disa = True and cert_start_date = "__ __ ____") then
        script_end_procedure("Client is certified disabled through SSA. Both SSA disability dates and PSN dates need to be listed on STAT/DISA. The script will now end.")
    Else
        DISA_start = replace(cert_start_date, " ", "/")
        If DISA_start = "__/__/____" then DISA_start = ""

        DISA_end = replace(cert_end_date, " ", "/")
        If DISA_end = "__/__/____" then DISA_end = ""
    End if
Else
    EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
	EMReadScreen disa_end_date, 10, 6, 69
	IF disa_start_date <> "__/__/____" then disa_start = Replace(disa_start_date," ","/")		'cleans up DISA dates
	If disa_end_date <> "__ __ ____" then disa_end = Replace(disa_end_date," ","/")

    DISA_start = replace(disa_start_date, " ", "/")
    If DISA_start = "__/__/____" then DISA_start = ""

    DISA_end = replace(disa_end_date, " ", "/")
    If DISA_end = "__/__/____" then DISA_end = ""
End if

If cdate(DISA_start) <= cdate("02/01/2018") then DISA_start = "02/01/2018"

'logic to ensure that the disa end date extends through the end of the month if necessary.
If disa_end <> "" then
    next_month = DateAdd("M", 1, DISA_end)
    next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
    DISA_end = dateadd("d", -1, next_month)
End if

'----------------------------------------------------------------------------------------------------BUSI and JOBS panels
CSR_required = ""   'Value that will be established as True or False based on if someone is working or not.

Call navigate_to_MAXIS_screen("STAT", "JOBS")
EmWriteScreen "01", 20, 76
Call write_value_and_transmit("01", 20, 79)
EMReadScreen total_panels, 1, 2, 78
If total_panels = "0" then
    CSR_required = FALSE
Else
    Do
        EmReadscreen JOBS_end_date, 8, 9, 49
        If JOBS_end_date = "__ __ __" then
            CSR_required = True
            exit do
        ELSE
            CSR_required = FALSE
            transmit
        End if
        EmReadscreen error_check, 5, 24, 2
    Loop until error_check = "ENTER"
End if

If CSR_requied <> True then
    Call navigate_to_MAXIS_screen("STAT", "BUSI")
    EmWriteScreen "01", 20, 76
    Call write_value_and_transmit("01", 20, 79)
    EMReadScreen total_panels, 1, 2, 78
    If total_panels = "0" then
        CSR_required = FALSE
    Else
        Do
            EmReadscreen BUSI_end_date, 8, 5, 72
            If BUSI_end_date = "__ __ __" then
                CSR_required = True
                exit do
            ELSE
                CSR_required = FALSE
                transmit
            End if
            EmReadscreen error_check, 5, 24, 2
        Loop until error_check = "ENTER"
    End if
End if
'----------------------------------------------------------------------------------------------------REVW panel
Call navigate_to_MAXIS_screen("STAT", "REVW")
Call write_value_and_transmit("X", 5, 35)

If CSR_required = True then
    EmReadscreen SR_month, 2, 9, 26
    EmReadscreen SR_year, 2, 9, 32
    If SR_month = "__" then script_end_procedure("A CSR is required for this case due to earned income. Please update the case, and run the script again if needed. The script will now end.")
End if

PF3 'back to stat/revw screen
EmReadscreen next_revw_month, 2, 9, 37
EmReadscreen next_revw_day, 2, 9, 40
EmReadscreen next_revw_year, 2, 9, 43

next_revw_date = next_revw_month & "/" & next_revw_day & "/" & next_revw_year
revw_end = dateadd("d", -1, next_revw_date)
IF CSR_required = true then
    revw_start = dateadd("M", - 6, next_revw_date)
else
    revw_start = dateadd("M", - 12, next_revw_date)
End if

If cdate(revw_start) <= cdate("02/01/2018") then revw_start = "02/01/2018"

'----------------------------------------------------------------------------------------------------SSRT: ensuring that a panel exists, and the FACI dates match.
 Call navigate_to_MAXIS_screen ("STAT", "SSRT")
 Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01.
 Call write_value_and_transmit ("01", 20, 79)    'Ensuring we're on the 1st panel

EMReadScreen SSRT_total_check, 1, 2, 78
If SSRT_total_check = "0" then
	script_end_procedure("SSRT panel needs to be created. The script will now end.")
Elseif SSRT_total_check = "1" then
    SSRT_found = True
Else
    Do
        confirm_SSRT = msgbox("Is this the facility/vendor you'd like to create an agreement for? Press NO to check next facility. Press YES to continue.", vbYesNoCancel + vbQuestion, "More than one SSRT panel exists.")
	    If confirm_SSRT = vbCancel then script_end_procedure("You have pressed Cancel. The script will now end.")
        If confirm_SSRT = vbYes then
            SSRT_found = True
            exit do
        End if
        If confirm_SSRT = vbNo then
            SSRT_found = False
            transmit
            EmReadscreen last_panel, 5, 24, 2
        End if
    Loop until last_panel = "ENTER"
    If SSRT_found = False then script_end_procedure("All facility/vendors have reviewed without being selected. The script will now end.")
End if

'Trying to find a suggested date based on the SSRT panel
EMReadScreen SSRT_vendor_number, 8, 5, 43		'Enters vendor number
EmReadscreen SSRT_vendor_name, 30, 6, 43
SSRT_vendor_name = replace(SSRT_vendor_name, "_", "")
EMReadScreen NPI_number, 10, 7, 43

If trim(NPI_number) = "" then script_end_procedure("No NPI number on SSRT panel. Agreement cannot be loaded into MMIS. Please report this NPI number to DHS. The script will now end.")
If instr(SSRT_vendor_name, "ANDREW RESIDENCE") then script_end_procedure("Andrew Residence facilities do not get loaded into MMIS. The script will now end.")

current_faci = false
row = 14
Do
    EMReadScreen ssrt_out_date, 10, row, 71
    EMReadScreen ssrt_in_date, 10, row, 47
    If ssrt_out_date = "__ __ ____" then
        If ssrt_in_date = "__ __ ____" then
            current_faci = False
            row = row - 1
        else
            current_faci = True
            Exit do
        End if
    Else
        current_faci = true
        exit do
    End if
    If row = 9 then
        transmit
        row = 14
    End if
Loop until row = 9

SSRT_start = replace(ssrt_in_date, " ", "/")
If SSRT_start = "__/__/____" then SSRT_start = ""
If cdate(SSRT_start) <= cdate("02/01/2018") then SSRT_start = "02/01/2018"

SSRT_end = replace(ssrt_out_date, " ", "/")
If SSRT_end = "__/__/____" then SSRT_end = ""

'----------------------------------------------------------------------------------------------------MEMB and ADDR panels
Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadScreen client_PMI, 8, 4, 46
client_PMI = trim(client_PMI)
client_PMI = right("00000000" & client_pmi, 8)

EMReadScreen client_DOB, 10, 8, 42
client_DOB = replace(client_DOB, " ", "")

Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
If mail_line_one = "" then
	addr_line_01 = resi_line_one
	addr_line_02 = resi_line_two
	city_line = resi_city
	state_line = resi_state
	zip_line = resi_zip
Else
	addr_line_01 = mail_line_one
	addr_line_02 = mail_line_two
	city_line = mail_city
	state_line = mail_state
	zip_line = mail_zip
End if

'----------------------------------------------------------------------------------------------------FACI panel
Call navigate_to_MAXIS_screen("STAT", "FACI")
Call write_value_and_transmit ("01", 20, 76) 	'For member 01 - All GRH cases should be for member 01.
Call write_value_and_transmit ("01", 20, 79)    'Ensuring we're on the 1st panel

'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
EMReadScreen FACI_total_check, 1, 2, 78
If FACI_total_check = "0" then script_end_procedure("Case does not have a FACI panel. The script will now end.")
'Matching the FACI panel vendor number to the SSRT panel vendor number
faci_found = False
Do
    EMReadScreen faci_vendor_number, 8, 5, 43		'Enters vendor number
    If faci_vendor_number = SSRT_vendor_number then
        faci_found = true
        'Gathering approval county information
        EMReadScreen approval_county, 2, 12, 71
        approval_county = "0" & approval_county
        exit do	'when the correct
    else
        faci_found = False
    End if
    transmit
    EMReadScreen last_panel, 5, 24, 2
Loop until last_panel = "ENTER"	'This means that there are no other faci panels

If (faci_found = true and approval_county = "__") then script_end_procedure("Please fill in the 'Approval Cty' field on the FACI panel.")
If faci_found = False then script_end_procedure("FACI panel could not be found for the SSRT panel vendor. The script will now end.")

'----------------------------------------------------------------------------------------------------VNDS/VND2
Call Navigate_to_MAXIS_screen("MONY", "VNDS")
Call write_value_and_transmit(SSRT_vendor_number, 4, 59)
Call write_value_and_transmit("VND2", 20, 70)
EMReadScreen VND2_check, 4, 2, 54
If VND2_check <> "VND2" then script_end_procedure("Unable to find MONY/VND2 panel. The script will now end.")
EMReadScreen service_rate, 8, 16, 68		'Reading the service rate to input into MMIS
If IsNumeric(service_rate) = False then EMReadScreen service_rate, 8, 15, 72        'Handling for vendors with Rate 3 information
service_rate = replace(service_rate, ".", "")	'removing the period for input into MMIS
service_rate = trim(service_rate)

'----------------------------------------------------------------------------------------------------ELIG/GRH
'Trimming the vendor number of the preceding 0's since ELIG/GRH doesn't show the 0's.
If left(SSRT_vendor_number, 1) = "0" then
    Do
        SSRT_vendor_number = right(SSRT_vendor_number, len(SSRT_vendor_number) - 1)
    Loop until left(SSRT_vendor_number, 1) <> "0"
End if

Call Navigate_to_MAXIS_screen("ELIG", "GRH ")
EMReadScreen no_grh, 10, 24, 2		'NO GRH version means no conversion to MMIS will take place
If no_grh = "NO VERSION" then script_end_procedure("There are no GRH eligibility results. Please review. The script will now end.")

Call write_value_and_transmit("99", 20, 79)
'This brings up the FS versions of eligibility results to search for approved versions
status_row = 7
Do
	EMReadScreen app_status, 8, status_row, 50
	If trim(app_status) = "" then script_end_procedure("There are no GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". The script will now end.")
	If app_status = "UNAPPROV" Then status_row = status_row + 1
	IF app_status = "APPROVED" then
		EMReadScreen vers_number, 1, status_row, 23
		Call write_value_and_transmit(vers_number, 18, 54)
		exit do
 	End if
Loop until app_status = "APPROVED" or trim(app_status) = ""

If app_status <> "APPROVED" then script_end_procedure("There are no approved GRH eligibility results for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". The script will now end.")

'----------------------------------------------------------------------------------------------------ELIG/GRFB
Call write_value_and_transmit("GRFB", 20, 71)
Call write_value_and_transmit("x", 11, 3)
'Ensuring a rate 2 is found. If none or more than one are found, MMIS will not be updated.
row = 15
Do
    EMReadScreen rate_two_check, 8, row, 8
    rate_two_check = Trim(rate_two_check)
    If rate_two_check = SSRT_vendor_number then
        exit do
    else
        row = row + 1
    End if
Loop until row = 20

If rate_two_check = "" then script_end_procedure("GRH eligibility doesn't reflect Rate 2 vendor information, or the SSRT vendor number did not match ELIG/GRFB vendor number. The script will now end.")
PF3' out of ELIG/GRFB

'----------------------------------------------------------------------------------------------------Main selection dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 170, "Select the SSR start and end dates for "  & SSRT_vendor_name
  CheckBox 60, 25, 65, 10, DISA_start, disa_start_checkbox
  CheckBox 150, 25, 65, 10, DISA_end, disa_end_checkbox
  CheckBox 60, 45, 65, 10, revw_start, revw_start_checkbox
  CheckBox 150, 45, 65, 10, revw_end, revw_end_checkbox
  CheckBox 60, 65, 65, 10, SSRT_start, SSRT_start_checkbox
  CheckBox 150, 65, 65, 10, SSRT_end, SSRT_end_checkbox
  EditBox 60, 85, 55, 15, custom_start
  EditBox 150, 85, 55, 15, custom_end
  EditBox 85, 110, 190, 15, custom_dates_explained
  EditBox 85, 130, 190, 15, other_notes
  EditBox 85, 150, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 150, 40, 15
    CancelButton 235, 150, 40, 15
    PushButton 15, 25, 30, 10, "DISA", DISA_Button
    PushButton 15, 45, 30, 10, "REVW", REVW_button
    PushButton 15, 65, 30, 10, "SSRT", SSRT_Button
  Text 5, 115, 75, 10, "Explain custom dates:"
  Text 5, 90, 45, 10, "Custom date:"
  Text 5, 135, 75, 10, "Other SSR/GRH notes:"
  GroupBox 145, 10, 85, 95, "Select the SSR end date"
  Text 20, 155, 60, 10, "Worker signature:"
  GroupBox 50, 10, 85, 95, "Select the SSR start date"
  GroupBox 235, 10, 40, 75, "Navigation"
  ButtonGroup ButtonPressed
    PushButton 240, 40, 30, 10, "JOBS", JOBS_button
    PushButton 240, 25, 30, 10, "FACI", FACI_button
    PushButton 240, 55, 30, 10, "MAXIS", MAXIS_button
    PushButton 240, 70, 30, 10, "MMIS", MMIS_button
EndDialog

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
    DO
        DO
            dialog Dialog1				'main dialog
            cancel_confirmation
            'Navigation button handling
            MAXIS_dialog_navigation
            If ButtonPressed = MAXIS_button then Call navigate_to_MAXIS(maxis_mode)  'Function to navigate back to MAXIS
            If ButtonPressed = MMIS_button then
                Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
                Call MMIS_panel_confirmation("AKEY", 51)         'ensuring we are on the right MMIS screen
                Call write_value_and_transmit(client_PMI, 10, 36)
            End if

            start_date = "" 'revaluing the variables for the start_date date
            end_date = ""   'revaluing the variables for the end date
            custom_date = ""
            total_units = ""
            If disa_start_checkbox = checked then start_date = start_date & disa_start
            If revw_start_checkbox = checked then start_date = start_date & revw_start
            If SSRT_start_checkbox = checked then start_date = start_date & SSRT_start
            If trim(custom_start) <> "" then
                start_date = start_date & custom_start
                custom_date = true
            End if

            If Disa_end_checkbox = checked then end_date = end_date & DISA_end
            If revw_end_checkbox = checked then end_date = end_date & revw_end
            If SSRT_end_checkbox = checked then end_date = end_date & SSRT_end
            If trim(custom_end) <> "" then
                end_date = end_date & custom_end
                custom_date = true
            End if
        Loop until ButtonPressed = -1

        err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
        If trim(start_date) = "" or IsDate(start_date) = false THEN err_msg = err_msg & vbCr & "Select/enter one valid start date."		'mandatory field
        IF trim(end_date) = "" or IsDate(end_date) = false THEN err_msg = err_msg & vbCr & "Select/enter one valid end date."		'mandatory field
        'If total_units > 365 THEN err_msg = err_msg & vbCr & "You cannot enter an agreement for more than 365 days. Select a new start and/or end dates."   ' Cannot be over 365 days.
        If (custom_date = True and trim(custom_dates_explained) = "") THEN err_msg = err_msg & vbCr & "Explain the reason for selecting custom dates."		'mandatory field
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "Enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'------------------------------------------------------------------------------------------------------calcuationas and conversion for MMIS
total_units = datediff("d", start_date, end_date) + 1   'Determining the total units to enter into MMIS.
MAXIS_agree_period = start_date & "-" & end_date

start_mo =  right("0" &  DatePart("m",    start_date), 2)
start_day = right("0" &  DatePart("d",    start_date), 2)
start_yr =  right(       DatePart("yyyy", start_date), 2)

output_start_date = start_mo & start_day & start_yr

end_mo =  right("0" &  DatePart("m",    end_date), 2)
end_day = right("0" &  DatePart("d",    end_date), 2)
end_yr =  right(       DatePart("yyyy", end_date), 2)

output_end_date = end_mo & end_day & end_yr

'----------------------------------------------------------------------------------------------------MMIS portion of the script
Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
Call MMIS_panel_confirmation("AKEY", 51)         'ensuring we are on the right MMIS screen

EmWriteScreen client_PMI, 10, 36
Call write_value_and_transmit("C", 3, 22)	'Checking to make sure that more than one agreement is not listed by trying to change (C) the information for the PMI selected.
EMReadScreen active_agreement, 12, 24, 2

If active_agreement = "NO DOCUMENTS" then
    duplicate_agreement = False 'no agreements exists in MMIS
else
	EMReadScreen AGMT_status, 31, 3, 19
    AGMT_status = trim(AGMT_status)
    If AGMT_status = "START DT:        END DT:" then
        row = 6
        Do
            EMReadScreen agreement_status, 1, row, 60
            EMReadScreen ASEL_start_date, 6, row, 63
            EmReadscreen ASEL_end_date, 6, row, 70
            ASEL_period = ASEL_start_date & "-" & ASEL_end_date
            output_period = output_start_date & "-" & output_end_date
            If agreement_status = "A" then
                If ASEL_period = output_period then
                    duplicate_agreement = True
                    script_end_procedure("An approved agreement already exists for the time frame selected. Please review the case. The script will now end.")
                Else
                    duplicate_agreement = False
                    row = row + 1
                End if
            ElseIf agreement_status = "D" then
                duplicate_agreement = False
                row = row + 1
            Else
                duplicate_agreement = False
            End if
        Loop until trim(agreement_status) = ""
    Else
        EMReadScreen agreement_status, 8, 3, 19
        If agreement_status = "APPROVED" then
            EMReadScreen ASA1_start_date, 6, row, 63
            EmReadscreen ASA1_end_date, 6, row, 70
            ASA1_period = ASA1_start_date & "-" & ASA1_end_date
            output_period = output_start_date & "-" & output_end_date

            If ASA1_period = output_period then
                duplicate_agreement = True
                script_end_procedure("An approved agreement already exists for the time frame selected. Please review the case. The script will now end.")
            Else
                duplicate_agreement = false
            End if
        Else
            duplicate_agreement = False
        End if
    End if
    PF6 'back to AKEY screen
End if

If duplicate_agreement = true then script_end_procedure("It appears an approved agreement already exists. Please review the case. The script will now end.")

If duplicate_agreement = False then
    Call clear_line_of_text(10, 36) 	'clears out the PMI number. Cannot add new agreement with PMI listed on AKEY.
    EmWriteScreen "A", 3, 22					'Selects the action code (A)
    EmWriteScreen "T", 3, 71					'Selecs the service agreement option (T)
    Call write_value_and_transmit("2", 7, 77)	'Enters the agreement type and transmits

    '----------------------------------------------------------------------------------------------------ASA1 screen
    Call MMIS_panel_confirmation("ASA1", 51)         'ensuring we are on the right MMIS screen

    EmWriteScreen output_start_date, 4, 64				'Start date
    EmWriteScreen output_end_date, 4, 71				'End date
    EmWriteScreen client_PMI, 8, 64						'Enters the client's PMI
    EmWriteScreen client_DOB, 9, 19						'Enters the client's DOB
    EmWriteScreen approval_county, 11, 19				'Enters 3 digit CO of SVC
    EmWriteScreen approval_county, 11, 39				'Enters 3 digit CO of RES
    Call write_value_and_transmit(approval_county, 11, 64)	'Enters 3 digit CO of FIN RESP and transmits

    Call MMIS_panel_confirmation("ASA2", 51)         'ensuring we are on the right MMIS screen
    transmit 	'no action required on ASA2
    '----------------------------------------------------------------------------------------------------ASA3 screen
    Call MMIS_panel_confirmation("ASA3", 51)         'ensuring we are on the right MMIS screen
    EMWriteScreen "H0043", 7, 36
    EMWriteScreen "U5", 7, 44
    EmWriteScreen output_start_date, 8, 60
    EmWriteScreen output_end_date, 8, 67
    EMWriteScreen service_rate, 9, 20			'Enters service rate from VND2
    EMWriteScreen total_units, 9, 60

    Call write_value_and_transmit(NPI_number, 10, 20)	'Enters the NPI number then transmits
    Emreadscreen NPI_issue, 26, 24, 1
    If NPI_issue = "CORRECT HIGHLIGHTED FIELDS" then
    	Update_MMIS = False
    	script_end_procedure("Issue with NPI# in MMIS. Please review case/report issue to the Quality Improvement Team. The script will now end.")
    	Call clear_line_of_text(10, 20) 	'clears out the NPI number so that the rest of the information can be saved.
    	PF3
    else
        '----------------------------------------------------------------------------------------------------PPOP screen handling
        EMReadScreen PPOP_check, 4, 1, 52
        If PPOP_check = "PPOP" then
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 180, 90, "PPOP screen - Choose Facility"
                ButtonGroup ButtonPressed
                OkButton 65, 70, 50, 15
                CancelButton 120, 70, 50, 15
                Text 5, 5, 170, 35, "Please select the correct facility name/address from the list in PPOP by putting a 'X' next to the name. DO NOT TRANSMIT. Press OK when ready. Press CANCEL to stop the script."
                Text 5, 45, 175, 20, "* Provider types for GRH must be '18/H COMM PRV' and the status must be '1 ACTIVE.'"
            EndDialog
            Do
                dialog Dialog1
                cancel_confirmation
            Loop until ButtonPressed = -1
			EMReadScreen PPOP_check, 4, 1, 52
            If PPOP_check = "PPOP" then transmit     'to exit PPOP
            If PPOP_check = "SA3 " then transmit    'to navigate to ACF1 - this is the partial screen check for ASA3
            transmit ' to next available screen (does not need to be updated)
            Call write_value_and_transmit("ACF1", 1, 51)
        End if

        '----------------------------------------------------------------------------------------------------ACF1 screen
        Call MMIS_panel_confirmation("ACF1", 51)         'ensuring we are on the right MMIS screen
        EmWriteScreen addr_line_01, 5, 8	'enters the clients address
        EmWriteScreen addr_line_02, 5, 37
        EmWriteScreen city_line, 6, 8
        EmWriteScreen state_line, 6, 34
        EmWriteScreen zip_line, 6, 42
        Call write_value_and_transmit("ASA1", 1, 8)		'direct navigating to ASA1

        '----------------------------------------------------------------------------------------------------ASA1 screen
        Call MMIS_panel_confirmation("ASA1", 51)         'ensuring we are on the right MMIS screen
         PF9 								'triggering stat edits
        EmreadScreen error_codes, 79, 20, 2	'checking for stat edits
        If trim(error_codes) <> "00 140  4          01 140  4" then
        	script_end_procedure("MMIS stat edits exist. Edit codes are: " & error_codes & vbcr & "PF3 to save what's been updated in MMIS, and follow up on the error codes. The script will now end.")
        else
        	EMWriteScreen "A", 3, 17						'Updating the AMT type/STAT to A for approved
        	Call write_value_and_transmit("ASA3", 1, 8)		'direct navigating to ASA3
        	Call MMIS_panel_confirmation("ASA3", 51)         'ensuring we are on the right MMIS screen
        	EMWriteScreen "A", 12, 19						'Updating the STAT CD/DATE to A for approved
        	Update_MMIS = true
            PF3 '	to save changes

            Call MMIS_panel_confirmation("AKEY", 51)         'ensuring we are on the right MMIS screen
            EMReadScreen authorization_number, 13, 9, 36
            authorization_number = trim(authorization_number)
            EMReadscreen approval_message, 16, 24, 2
        End if
    End if
End if

'----------------------------------------------------------------------------------------------------Back to MAXIS & CASE/NOTE
If Update_MMIS = True then
    Call navigate_to_MAXIS(maxis_mode)  'Function to navigate back to MAXIS
    Call check_for_MAXIS(False)

    If disa_start_checkbox = checked then start_date_source = ", PSN start date."
    If revw_start_checkbox = checked then start_date_source = ", start of certification period."
    If SSRT_start_checkbox = checked then start_date_source = ", SSRT start date."

    If Disa_end_checkbox = checked then end_date_source = ", PSN end date."
    If revw_end_checkbox = checked then end_date_source = ", end of certification period."
    If SSRT_end_checkbox = checked then end_date_source = ", SSRT end date."

    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9
    Call write_variable_in_CASE_NOTE("GRH Rate 2 SSR added to MMIS for " & SSRT_vendor_name)
    Call write_bullet_and_variable_in_CASE_NOTE("NPI #", npi_number)
    Call write_bullet_and_variable_in_CASE_NOTE("MMIS authorization number", authorization_number)
    Call write_variable_in_CASE_NOTE("* SSR start date: " & start_date & start_date_source)   'Hard coded for now
    Call write_variable_in_CASE_NOTE("* SSR end date: " & end_date & end_date_source)
    Call write_bullet_and_variable_in_CASE_NOTE("Explanation of custom date", custom_dates_explained)
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
    PF3
End if

script_end_procedure("Success! Your case has been updated in MMIS and case noted in MAXIS.")
