name_of_script = "UTILITIES - QI AVS REQUEST.vbs"
start_time = timer
STATS_counter = 0                     	'sets the stats counter at zero
STATS_manualtime = 300               	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "M"       		'C is for each CASE
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
call changelog_update("09/24/2021", "GitHub Issue #583 Updates made to ensure email has information went sent to QI", "MiKayla Handley, Hennepin County")
call changelog_update("09/08/2021", "Added date completed AVS form rec'd to dialog and reminder that completed AVS needs to be on file prior to submitting AVS request.", "Ilse Ferris, Hennepin County")
call changelog_update("09/30/2020", "Updated closing message.", "Ilse Ferris, Hennepin County")
call changelog_update("03/10/2020", "Initial version.", "MiKayla Handley, Hennepin County")

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
'Connecting to BlueZone
EMConnect ""
'Grabs the case number
CALL MAXIS_case_number_finder (MAXIS_case_number)
closing_message = "Request for Account Validation Service (AVS) email has been sent." 'setting up closing_message or possible additions later based on conditions
'----------------------------------------------------------------------------------------------------Initial dialog
appl_type = "Application"
applicant_type = "Applicant"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 211, 160, "AVS Request"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 185, 5, 20, 15, HH_size
  EditBox 155, 25, 50, 15, avs_form_date
  DropListBox 80, 60, 125, 15, "Select One:"+chr(9)+"Applicant"+chr(9)+"Spouse", applicant_type
  DropListBox 80, 80, 125, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Renewal", appl_type
  DropListBox 80, 100, 125, 15, "Select One:"+chr(9)+"BI-Brain Injury Waiver"+chr(9)+"BX-Blind"+chr(9)+"CA-Community Alt. Care"+chr(9)+"DD-Developmental Disa Waiver"+chr(9)+"DP-MA for Employed Pers w/ Disa"+chr(9)+"DX-Disability"+chr(9)+"EH-Emergency Medical Assistance"+chr(9)+"EW-Elderly Waiver"+chr(9)+"EX-65 and Older"+chr(9)+"LC-Long Term Care"+chr(9)+"MP-QMB SLMB Only"+chr(9)+"QI-QI"+chr(9)+"QW-QWD", MA_type
  DropListBox 80, 120, 125, 15, "Select One:"+chr(9)+"NA-No Spouse"+chr(9)+"YES"+chr(9)+"NO", spouse_deeming
  ButtonGroup ButtonPressed
    OkButton 110, 140, 45, 15
    CancelButton 160, 140, 45, 15
  Text 5, 65, 50, 10, "Applicant Type:"
  Text 5, 85, 65, 10, "Application Type:"
  Text 5, 105, 55, 10, "Request Type:"
  Text 5, 125, 35, 10, "Deeming:"
  Text 150, 10, 30, 10, "HH Size:"
  Text 5, 30, 120, 10, "Date Completed AVS Form Received:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 45, 200, 10, "AVS form must be complete and valid to submit AVS Request."
EndDialog

DO
    DO
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If HH_size = "" or IsNumeric(HH_size) = False then err_msg = err_msg & vbNewLine & "* Please enter a valid household composition size."
        If trim(avs_form_date) = "" or isdate(avs_form_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date the completed AVS form was received in the agency."
		IF applicant_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the applicant type."
		IF appl_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the application type."
		IF MA_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the MA request type."
		IF spouse_deeming = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select if the spouse is deeming."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password_without_transmit(are_we_passworded_out)
Loop until are_we_passworded_out = false
CALL check_for_MAXIS(False)
CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv) 'navigating to stat prog to gather the application information
IF is_this_priv = TRUE THEN script_end_procedure("PRIV case, cannot access/update. The script will now end.")

EMReadScreen application_date, 8, 12, 33 'Reading the HC app date from PROG
application_date = replace(application_date, " ", "/")
IF application_date = "__/__/__"  THEN script_end_procedure("*** No application date ***" & vbNewLine & "Need to have pending or active HC care to request AVS.")

CALL HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.

CALL navigate_to_MAXIS_screen("STAT", "MEMB") 'navigating to stat memb to gather the ref number and name.

DO
    CALL HH_member_custom_dialog(HH_member_array)
    IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
LOOP UNTIL uBound(HH_member_array) <> -1

CALL get_county_code
EMReadscreen current_county, 4, 21, 21
If lcase(current_county) <> worker_county_code THEN script_end_procedure("Out of County case, cannot access/update. The script will now end.")

'Establishing array
avs_membs = 0       'incrementor for array
DIM avs_members_array()  'Declaring the array this is what this list is
ReDim avs_members_array(phone_type_three_const, 0)  'Resizing the array 'redimmed to the size of the last constant  'that ,list is going to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const maxis_case_number_const  	 	= 0 '=  Maxis'
const member_number_const   		= 1 '=  Member Number
const client_first_name_const       = 2 '=  First Name MEMB
const client_last_name_const        = 3 '=  Last Name MEMB
const client_mid_name_const    	    = 4 '=  Middle initial MEMB
const client_DOB_const   		    = 5 '=  Date of Birth MEMB
const client_ssn_const		        = 6 '=  SSN
const client_age_const	            = 7 '=  age MEMB
const client_sex_const			    = 8 '=  client sex
const addr_eff_date_const	  		= 9 '=	addr_eff_date
const resi_line_one_const	  		= 10'= 	resi_line_one
const resi_line_two_const	   		= 11 '= resi_line_two
const resi_city_const				= 12'= 	resi_city
const resi_state_const	     		= 13'= 	resi_state
const resi_zip_const     			= 14'= 	resi_zip
const resi_county_const     		= 15'= 	resi_county
const verif_const	  				= 16'= 	verif
const homeless_const     			= 17 '= homeless
const ind_reservation_const  		= 18 '= ind_reservation
const living_sit_const	     		= 19 '= living_sit
const res_name_const    			= 20'= 	res_name
const mail_line_one_const    		= 21'= 	mail_line_one
const mail_line_two_const     		= 22'=  mail_line_two
const mail_city_const   			= 23'= 	mail_city
const mail_state_const    			= 24'= 	mail_state
const mail_zip_const    			= 25'= 	mail_zip
const phone_numb_one_const    		= 26'= 	phone_numb_one
const phone_type_one_const    		= 27'= 	phone_type_one
const phone_numb_two_const     		= 28'= 	phone_numb_two
const phone_type_two_const  	   	= 29'= 	phone_type_two
const phone_numb_three_const     	= 30'= 	phone_numb_three
const phone_type_three_const    	= 31'= 	phone_type_three

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
FOR EACH person IN HH_member_array
    CALL write_value_and_transmit(person, 20, 76) 'reads the reference number, last name, first name, and THEN puts it into an array YOU HAVENT defined the avs_members_array yet
    EMReadscreen ref_nbr, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen MEMB_number, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EMReadscreen client_SSN, 11, 7, 42
    If client_ssn = "___ __ ____" then client_ssn = ""
    last_name = trim(replace(last_name, "_", "")) & " "
    first_name = trim(replace(first_name, "_", "")) & " "
    mid_initial = replace(mid_initial, "_", "")
    EMReadScreen client_age, 2, 8, 76
    IF client_age = "  " THEN client_age = 0
    client_age = client_age * 1
	EMReadScreen client_sex, 1, 9, 42
    ReDim Preserve avs_members_array(phone_type_three_const, avs_membs)  'redimmed to the size of the last constant
    avs_members_array(member_number_const,     avs_membs) = ref_nbr
    avs_members_array(client_first_name_const, avs_membs) = first_name
    avs_members_array(client_last_name_const,  avs_membs) = last_name
    avs_members_array(client_mid_name_const,   avs_membs) = mid_initial
    avs_members_array(client_DOB_const,        avs_membs) = client_DOB
    avs_members_array(client_ssn_const,        avs_membs) = client_SSN
    avs_members_array(client_age_const,        avs_membs) = client_age
	avs_members_array(client_sex_const,        avs_membs) = client_sex
    avs_membs = avs_membs + 1 ' can only be used because we havent reset or redefined this incrementor'
	STATS_counter = STATS_counter + 1
NEXT

CALL navigate_to_MAXIS_screen("STAT", "MEMI")
EMReadScreen marital_status, 1, 7, 40
EMReadScreen spouse_ref_nbr, 02, 09, 49
spouse_ref_nbr = replace(spouse_ref_nbr, "_", "")
IF marital_status = "M" and spouse_ref_nbr <> "" THEN client_married = TRUE
IF spouse_deeming = "YES" and spouse_ref_nbr = "" THEN
	BeginDialog Dialog1, 0, 0, 176, 160, "Spouse not found on MEMB"
      EditBox 55, 5, 115, 15, spouse_first_name
      EditBox 55, 25, 115, 15, spouse_last_name
      EditBox 55, 45, 35, 15, spouse_mid_name
      EditBox 120, 45, 50, 15, spouse_SSN_number
      EditBox 55, 65, 55, 15, spouse_DOB
      EditBox 150, 65, 20, 15, spouse_age
      DropListBox 55, 85, 55, 15, "Select One:"+chr(9)+"Female"+chr(9)+"Male"+chr(9)+"Unknown"+chr(9)+"Undetermined", spouse_gender_dropdown
      EditBox 5, 120, 165, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 75, 140, 45, 15
        CancelButton 125, 140, 45, 15
      Text 5, 10, 40, 10, "First Name:"
      Text 5, 30, 40, 10, "Last Name: "
      Text 5, 50, 50, 10, " Middle Initial: "
      Text 100, 50, 20, 10, "SSN: "
      Text 5, 70, 45, 10, "Date of Birth: "
      Text 130, 70, 15, 10, "Age: "
      Text 5, 105, 160, 10, "Please explain why they are not listed in maxis: "
      Text 5, 90, 30, 10, "Gender: "
    EndDialog

	DO
	 	DO
	 		err_msg = ""
	 		Dialog Dialog1
	 		cancel_without_confirmation
	 		If spouse_first_name = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's first name."
	        If spouse_last_name = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's last name."
	        If spouse_SSN_number = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's social security number."
	        If spouse_DOB = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's date of birth."
	        If spouse_gender_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "Please select the spouse's gender."
	        If other_notes = "" then err_msg = err_msg & vbNewLine & "Please enter the reason this client is not listed in MAXIS."
	 		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	 	LOOP UNTIL err_msg = ""
	 	CALL check_for_password_without_transmit(are_we_passworded_out)
	Loop until are_we_passworded_out = false

	ReDim Preserve avs_members_array(phone_type_three_const, avs_membs)  'redimmed to the size of the last constant
    avs_members_array(member_number_const,     avs_membs) = spouse_ref_nbr
    avs_members_array(client_first_name_const, avs_membs) = spouse_first_name
    avs_members_array(client_last_name_const,  avs_membs) = spouse_last_name
    avs_members_array(client_mid_name_const,   avs_membs) = spouse_mid_name
    avs_members_array(client_DOB_const,        avs_membs) = spouse_DOB
    avs_members_array(client_ssn_const,        avs_membs) = spouse_SSN_number
    avs_members_array(client_age_const,        avs_membs) = spouse_age
	avs_members_array(client_sex_const,        avs_membs) = spouse_gender_dropdown
	client_married = TRUE
END IF
' CALL read_ADDR_panel(addr_eff_date, line_one, line_two, city, state, zip, county, verif, homeless, ind_reservation, living_sit, res_name,                                mail_line_one, mail_line_two,                   mail_city, mail_state, mail_zip,                                  phone_one, type_one, phone_two, type_two, phone_three, type_three, updated_date)
Call access_ADDR_panel("READ", notes_on_address, line_one, line_two, resi_street_full, city, state, zip, county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

team_email = "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us"

FOR avs_membs = 0 to Ubound(avs_members_array, 2) 'start at the zero person and go to each of the selected people '
    member_info = member_info & "A signed AVS form was received for Member # " & avs_members_array(member_number_const, avs_membs) & vbNewLine & avs_members_array(client_first_name_const, avs_membs) & " " & avs_members_array(client_mid_name_const, avs_membs) & " " & avs_members_array(client_last_name_const, avs_membs)  &  vbCr & "DOB: " & avs_members_array(client_DOB_const,  avs_membs) & vbcr & "SSN of Resident: " & avs_members_array(client_ssn_const,  avs_membs) & vbcr & "Gender: " & avs_members_array(client_sex_const, avs_membs)
	member_info = member_info & vbNewLine & "AVS Form Received Date: " & avs_form_date & vbcr & "MA type: " & MA_type & vbcr & "HH size: " & HH_size & vbcr & "Applicant Type: " & applicant_type & vbcr & "Application Type: " & appl_type & vbNewLine & "Residential Address: " & vbNewLine & line_one & " " & line_two & vbcr & city & ", " & state & " " & zip
	If trim(mail_line_one) <> "" THEN member_info = member_info & "Mailing address: " & mail_line_one & vbcr & mail_line_two & vbcr & mail_city & vbcr & mail_state & vbcr & mail_zip & vbcr & phone_one & " Phone: " & type_one & " - " & phone_two & " - " & phone_three
	IF client_married = TRUE THEN member_info = member_info & vbNewLine & "Spouse: " & spouse_deeming & vbcr & "Spouse Member # " & avs_members_array(member_number_const, avs_membs) & vbcr & "Spouse First Name: " & avs_members_array(client_first_name_const, avs_membs) & vbcr & "Spouse Last Name: " & avs_members_array(client_last_name_const, avs_membs) & vbcr & "Spouse Social Security Number: " & avs_members_array(client_ssn_const,  avs_membs) & vbcr & "Spouse Gender: " & avs_members_array(client_sex_const, avs_membs) & vbcr & "Spouse Date of birth: " & avs_members_array(client_DOB_const, avs_membs) & " " & other_notes
NEXT

CALL find_user_name(the_person_running_the_script)' this is for the signature in the email'

'Creating the email
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachmentsend_email)
Call create_outlook_email(team_email, "", "AVS initial run requests case #" & MAXIS_case_number, member_info & vbNewLine & vbNewLine & "Submitted By: " & vbNewLine & the_person_running_the_script, "", TRUE)   'will create email, will send.

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/30/2021
'--Tab orders reviewed & confirmed----------------------------------------------08/30/2021
'--Mandatory fields all present & Reviewed--------------------------------------08/30/2021
'--All variables in dialog match mandatory fields-------------------------------08/30/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/09/21
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------08/30/2021
'--Out-of-County handling reviewed----------------------------------------------08/30/2021
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/09/21
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------09/09/21
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------08/30/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------08/30/2021
'--Comment Code-----------------------------------------------------------------09/09/21
'--Update Changelog for release/update------------------------------------------09/09/21
'--Remove testing message boxes-------------------------------------------------09/09/21
'--Remove testing code/unnecessary code-----------------------------------------09/09/21
'--Review/update SharePoint instructions----------------------------------------09/09/21
'--Review Best Practices using BZS page ----------------------------------------09/09/21
'--Review script information on SharePoint BZ Script List-----------------------09/09/21
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/09/21
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/09/21
'--Complete misc. documentation (if applicable)---------------------------------09/09/21
'--Update project team/issue contact (if applicable)----------------------------09/09/21
