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
call changelog_update("03/27/2024", "Updated the name of AVS to Asset Verification Service.", "Dave Courtright, Hennepin County")
call changelog_update("02/22/2024", "Enabled the Renewal option for submitting an AVS Request.", "Ilse Ferris, Hennepin County")
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("06/26/2023", "Disabled renewal option due to asset diregard through 05/31/2024.", "Ilse Ferris, Hennepin County")
call changelog_update("03/20/2023", "Added change in circumstance option, and updated email output and information gathering functionality.", "Ilse Ferris, Hennepin County")
call changelog_update("09/24/2021", "GitHub Issue #583 Updates made to ensure email has information went sent to QI", "MiKayla Handley, Hennepin County")
call changelog_update("09/08/2021", "Added date completed AVS form rec'd to dialog and reminder that completed AVS needs to be on file prior to submitting AVS request.", "Ilse Ferris, Hennepin County")
call changelog_update("09/30/2020", "Updated closing message.", "Ilse Ferris, Hennepin County")
call changelog_update("03/10/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""    'Connecting to BlueZone
CALL MAXIS_case_number_finder (MAXIS_case_number) 'Grabs the case number
Call check_for_MAXIS(FALSE)

'Initial Defaults
closing_message = "Request for Asset Verification Service (AVS) email has been sent." 'setting up closing_message or possible additions later based on conditions

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 211, 160, "AVS Request"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 185, 5, 20, 15, HH_size
  EditBox 155, 25, 50, 15, avs_form_date
  DropListBox 80, 60, 125, 15, "Select One:"+chr(9)+"Applicant"+chr(9)+"Spouse", applicant_type
  DropListBox 80, 80, 125, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Change In Basis"+chr(9)+"Renewal", HC_process
  DropListBox 80, 100, 125, 15, "Select One:"+chr(9)+"BI-Brain Injury Waiver"+chr(9)+"BX-Blind"+chr(9)+"CA-Community Alt. Care"+chr(9)+"CD - CADI Waiver"+chr(9)+"DD-Developmental Disa Waiver"+chr(9)+"DP-MA for Employed Pers w/ Disa"+chr(9)+"DX-Disability"+chr(9)+"EH-Emergency Medical Assistance"+chr(9)+"EW-Elderly Waiver"+chr(9)+"EX-65 and Older"+chr(9)+"LC-Long Term Care"+chr(9)+"MP-QMB SLMB Only"+chr(9)+"QI-QI"+chr(9)+"QW-QWD", MA_type
  DropListBox 80, 120, 125, 15, "Select One:"+chr(9)+"N/A - No Spouse"+chr(9)+"Yes"+chr(9)+"No", spouse_deeming
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
        Call validate_MAXIS_case_number(err_msg, "*")
		If trim(HH_size) = "" or IsNumeric(HH_size) = False then err_msg = err_msg & vbNewLine & "* Please enter a valid household composition size."
        If trim(avs_form_date) = "" or isdate(avs_form_date) = False then err_msg = err_msg & vbNewLine & "* Please enter the date the completed AVS form was received in the agency."
		IF applicant_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the applicant type."
		IF HC_process = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the application type."
		IF MA_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the MA request type."
		IF spouse_deeming = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select if the spouse is deeming."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call check_for_MAXIS(FALSE)
Call back_to_SELF

send_email = True 'Default to True, and enabling a no-send option for the script team and if the case is in training region. 
EMReadScreen MX_region, 10, 22, 48
IF trim(MX_region) = "TRAINING" THEN 
    send_email = FALSE
Elseif user_ID_for_validation = "CALO001" OR user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "MEGE001" OR user_ID_for_validation = "MARI001" OR user_ID_for_validation = "DACO003" Then 
    send_email = False
End if 

CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv) 'navigating to stat prog to gather the application information
IF is_this_priv = TRUE THEN script_end_procedure("PRIV case, cannot access/update. The script will now end.")

EMReadScreen application_date, 8, 12, 33 'Reading the HC app date from PROG
application_date = replace(application_date, " ", "/")
IF application_date = "__/__/__"  THEN script_end_procedure("*** No application date ***" & vbNewLine & "Need to have pending or active HC care to request AVS.")

CALL HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.
Call access_ADDR_panel("Read", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)    'reading ADDR panel information for later output in email/message box
CALL navigate_to_MAXIS_screen("STAT", "MEMB")

Do
    DO
        CALL HH_member_custom_dialog(HH_member_array)
        IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
    LOOP UNTIL uBound(HH_member_array) <> -1
    CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Establishing array
avs_membs = 0       'incrementor for array
DIM avs_members_array()  'Declaring the array this is what this list is
ReDim avs_members_array(marital_status_const, 0)  'Resizing the array 'redimmed to the size of the last constant  'that ,list is going to have 20 parameter but to start with there is only one parameter it gets complicated - grid'

'Creating constants to value the array elements this is why we create constants
const maxis_case_number_const  	 	= 0 '=  Maxis Case Number
const member_number_const   		= 1 '=  Member Number
const client_first_name_const       = 2 '=  First Name MEMB
const client_last_name_const        = 3 '=  Last Name MEMB
const client_mid_name_const    	    = 4 '=  Middle initial MEMB
const client_DOB_const   		    = 5 '=  Date of Birth MEMB
const client_ssn_const		        = 6 '=  SSN
const marital_status_const          = 7 '=  marital status

FOR EACH person IN HH_member_array
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    CALL write_value_and_transmit(person, 20, 76) 'reads the reference number, last name, first name, and THEN puts it into an array YOU HAVEN'T defined the avs_members_array yet

    EMReadscreen last_name, 25, 6, 30
    last_name = trim(replace(last_name, "_", ""))

    EMReadscreen first_name, 12, 6, 63
    first_name = trim(replace(first_name, "_", ""))

    EMReadscreen mid_initial, 1, 6, 79
    mid_initial = replace(mid_initial, "_", "")

    EMReadScreen client_DOB, 10, 8, 42

    EMReadscreen client_SSN, 11, 7, 42
    If client_ssn = "___ __ ____" then client_ssn = ""

    CALL navigate_to_MAXIS_screen("STAT", "MEMI")
    EmReadscreen martial_status, 1, 7, 40

    ReDim Preserve avs_members_array(marital_status_const, avs_membs)  'redimmed to the size of the last constant
    avs_members_array(member_number_const,     avs_membs) = person
    avs_members_array(client_first_name_const, avs_membs) = first_name
    avs_members_array(client_last_name_const,  avs_membs) = last_name
    avs_members_array(client_mid_name_const,   avs_membs) = mid_initial
    avs_members_array(client_DOB_const,        avs_membs) = client_DOB
    avs_members_array(client_ssn_const,        avs_membs) = client_SSN
    avs_members_array(marital_status_const,    avs_membs) = martial_status
    avs_membs = avs_membs + 1 ' can only be used because we haven't reset or redefined this incrementor'
	STATS_counter = STATS_counter + 1
NEXT

If avs_membs = 1 then
    'If user only selects one member and that member's martial status is M and the case is spouse deeming, the user will be asked to enter this information.
    If avs_members_array(marital_status_const, 0) = "M" then
        If spouse_deeming = "Yes" then
            manual_spouse_entry = True

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 176, 105, "Spouse not selected/available in MAXIS"
                EditBox 55, 5, 115, 15, spouse_first_name
                EditBox 55, 25, 115, 15, spouse_last_name
                EditBox 55, 45, 35, 15, spouse_mid_name
                EditBox 120, 45, 50, 15, spouse_SSN_number
                EditBox 55, 65, 55, 15, spouse_DOB
                ButtonGroup ButtonPressed
                    OkButton 75, 85, 45, 15
                    CancelButton 125, 85, 45, 15
                Text 10, 10, 40, 10, "First Name:"
                Text 10, 30, 40, 10, "Last Name:"
                Text 5, 50, 50, 10, " Middle Initial:"
                Text 100, 50, 20, 10, "SSN:"
                Text 5, 70, 45, 10, "Date of Birth:"
            EndDialog

	        DO
	         	DO
	         		err_msg = ""
	         		Dialog Dialog1
	         		cancel_confirmation
	         		If trim(spouse_first_name) = "" then err_msg = err_msg & vbNewLine & "* Enter the spouse's first name."
	                If trim(spouse_last_name) = "" then err_msg = err_msg & vbNewLine & "* Enter the spouse's last name."
	                If trim(spouse_SSN_number) = "" or len(spouse_SSN_number) < 9 then err_msg = err_msg & vbNewLine & "* Enter the spouse's 9-digit Social Security Number."
	                If trim(spouse_DOB) = "" or isdate(spouse_DOB) = False then err_msg = err_msg & vbNewLine & "* Enter the spouse's date of birth."
	         		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	         	LOOP UNTIL err_msg = ""
	         	CALL check_for_password(are_we_passworded_out)
	        Loop until are_we_passworded_out = FALSE
        End if
    End if
End if

'----------------------------------------------------------------------------------------------------Email/message box information
member_info = member_info & "A signed AVS form was received for case #" & MAXIS_case_number & vbcr & vbcr & _
"******Case Information******" & vbcr & _
"Application Date: " & application_date & vbcr & _
"AVS Form Received Date: " & avs_form_date & vbcr & _
"Basis of Eligibility: " & MA_type & vbcr & _
"HH size: " & HH_size & vbcr & _
"Applicant Type: " & applicant_type & vbcr & _
"Application Type: " & HC_process & vbcr & _
"Spouse Deeming?: " & spouse_deeming & vbcr

FOR avs_membs = 0 to Ubound(avs_members_array, 2) 'start at the zero person and go to each of the selected people '
    member_info = member_info & vbcr & _
    "Member Name: " & avs_members_array(client_first_name_const, avs_membs) & " " & avs_members_array(client_mid_name_const, avs_membs) & " " & avs_members_array(client_last_name_const, avs_membs)  & vbCr & _
    "Member #" & avs_members_array(member_number_const, avs_membs) & vbCr & _
    "DOB: " & avs_members_array(client_DOB_const, avs_membs) & vbcr & _
    "SSN: " & avs_members_array(client_ssn_const, avs_membs) & vbcr & _
    "MEMI Marital Status: " & avs_members_array(marital_status_const, avs_membs) & vbcr
Next

IF manual_spouse_entry = True then
    member_info = member_info & vbcr & "Manually Entered Spouse Information (Not in MAXIS):" & vbcr & _
    "* Spouse Name: " & spouse_first_name & " " & spouse_mid_name & " " & spouse_last_name & vbcr & _
    "* Spouse SSN: " & spouse_SSN_number & vbcr & _
    "* Spouse DOB: " & spouse_DOB & vbcr
End if

member_info = member_info & vbcr & "******Address Info******" & vbcr & _
"--Residential Address--" & vbcr & _
"Line 1: " & resi_line_one & vbcr & _
"Line 2: " & resi_line_two & vbcr & _
resi_city & ", " & resi_state & " " & resi_zip

If trim(mail_line_one) <> "" THEN
    member_info = member_info & vbcr & vbcr & "--Mailing Address--" & vbcr & _
    "Line 1: " & mail_line_one & vbcr & _
    "Line 2: " & mail_line_two & vbcr & _
    mail_city & ", " & mail_state & " " & mail_zip
End if

CALL find_user_name(the_person_running_the_script)' this is for the signature in the email'

If send_email = False then 
    script_end_procedure("!!EMAIL HAS NOT BEEN SENT!!" & vbcr & vbcr & "AVS initial run requests case #" & MAXIS_case_number & vbcr & vbcr & "Member Info:" & member_info & vbCR & vbcr & "Submitted By: " & the_person_running_the_script)
Elseif send_email = TRUE THEN 
    Call create_outlook_email("", "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "", "AVS initial run requests case #" & MAXIS_case_number, 1, False, "", "", False, "", member_info & vbNewLine & vbNewLine & "Submitted By: " & vbNewLine & the_person_running_the_script, False, "", True)
    script_end_procedure_with_error_report(closing_message)
End if 

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------03/20/2023
'--Tab orders reviewed & confirmed----------------------------------------------03/20/2023
'--Mandatory fields all present & Reviewed--------------------------------------03/20/2023
'--All variables in dialog match mandatory fields-------------------------------03/20/2023
'Review dialog names for content and content fit in dialog----------------------03/20/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------03/20/2023------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------03/20/2023------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------03/20/2023------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used 03/20/2023------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------03/20/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------03/20/2023------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------03/20/2023
'--Out-of-County handling reviewed----------------------------------------------03/20/2023------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------03/20/2023
'--BULK - review output of statistics and run time/count (if applicable)--------03/20/2023------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------03/20/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------03/20/2023
'--Incrementors reviewed (if necessary)-----------------------------------------03/20/2023
'--Denomination reviewed -------------------------------------------------------03/20/2023
'--Script name reviewed---------------------------------------------------------03/20/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------03/20/2023------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------03/20/2023
'--comment Code-----------------------------------------------------------------03/20/2023
'--Update Changelog for release/update------------------------------------------03/20/2023
'--Remove testing message boxes-------------------------------------------------03/20/2023
'--Remove testing code/unnecessary code-----------------------------------------03/20/2023
'--Review/update SharePoint instructions----------------------------------------03/20/2023------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------03/20/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------03/20/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------03/20/2023
'--Complete misc. documentation (if applicable)---------------------------------03/20/2023
'--Update project team/issue contact (if applicable)----------------------------03/20/2023
