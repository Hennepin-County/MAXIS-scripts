name_of_script = "UTILITIES - QI AVS REQUEST.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("09/08/2021", "Added date completed AVS form rec'd to dialog and reminder that completed AVS needs to be on file prior to submitting AVS request.", "Ilse Ferris, Hennepin County")
call changelog_update("09/30/2020", "Updated closing message.", "Ilse Ferris, Hennepin County")
call changelog_update("03/10/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connecting to BlueZone
EMConnect ""
'----------------------------------------------------------------------------------------------------ABPS panel

Call check_for_MAXIS(False)
'TO DO PRIV handling and check on hh size'
CALL MAXIS_case_number_finder (MAXIS_case_number)
appl_type = "Application"
'MEMB_number = "01"
'checking for an active MAXIS session
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 216, 175, "AVS Request"
  EditBox 50, 5, 45, 15, MAXIS_case_number
  EditBox 130, 5, 20, 15, HH_size
  EditBox 185, 5, 20, 15, MEMB_number
  EditBox 155, 30, 50, 15, avs_form_date
  DropListBox 80, 60, 125, 15, "Select One:"+chr(9)+"Applicant"+chr(9)+"Spouse", applicant_type
  DropListBox 80, 80, 125, 15, "Select One:"+chr(9)+"Application", appl_type
  'DropListBox 80, 80, 125, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Renewal", appl_type
  DropListBox 80, 100, 125, 15, "Select One:"+chr(9)+"BI-Brain Injury Waiver"+chr(9)+"BX-Blind"+chr(9)+"CA-Community Alt. Care"+chr(9)+"DD-Developmental Disa Waiver"+chr(9)+"DP-MA for Employed Pers w/ Disa"+chr(9)+"DX-Disability"+chr(9)+"EH-Emergency Medical Assistance"+chr(9)+"EW-Elderly Waiver"+chr(9)+"EX-65 and Older"+chr(9)+"LC-Long Term Care"+chr(9)+"MP-QMB SLMB Only"+chr(9)+"QI-QI"+chr(9)+"QW-QWD", MA_type
  DropListBox 80, 120, 125, 15, "Select One:"+chr(9)+"NA-No Spouse"+chr(9)+"YES"+chr(9)+"NO", spouse_deeming
  ButtonGroup ButtonPressed
    OkButton 110, 140, 45, 15
    CancelButton 160, 140, 45, 15
  Text 30, 65, 50, 10, "Applicant type:"
  Text 25, 85, 55, 10, "Application type:"
  Text 35, 105, 45, 10, "Request type:"
  Text 45, 125, 35, 10, "Deeming:"
  Text 100, 10, 30, 10, "HH size:"
  Text 20, 35, 130, 10, "Received Date of Completed AVS Form*:"
  Text 5, 10, 45, 10, "Case number:"
  Text 155, 10, 30, 10, "Memb #:"
  Text 5, 160, 205, 10, "* AVS form must be complete and valid to submit AVS Request."
EndDialog

DO
    DO
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If IsNumeric(MEMB_number) = False or len(MEMB_number) <> 2 then err_msg = err_msg & vbNewLine & "* Please enter a valid 2 digit member number."
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

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
EMReadScreen err_msg, 7, 24, 02

'Checking for PRIV cases.
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
IF priv_check = "PRIV" THEN MsgBox "*** Privilege Case ***" & vbNewLine & "This case is privileged please request access." & vbNewLine

EMReadScreen application_date, 8, 12, 33 'Reading the app date from PROG
application_date = replace(application_date, " ", "/")
IF application_date = "__/__/__" THEN MsgBox "*** No application date ***" & vbNewLine & "Need to have pending or active HC care to request AVS." & vbNewLine

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
EMReadScreen client_first_name, 12, 6, 63
client_first_name = replace(client_first_name, "_", "")
client_first_name = trim(client_first_name)
EMReadScreen client_last_name, 25, 6, 30
client_last_name = replace(client_last_name, "_", "")
client_last_name = trim(client_last_name)
EMReadscreen client_SSN_number_read, 11, 7, 42
client_SSN_number_read = replace(client_SSN_number_read, " ", "")
EmReadScreen client_DOB, 10, 8, 42
client_DOB = replace(client_DOB, " ", "/")
EmReadScreen client_gender, 1, 9, 42

IF spouse_deeming = "YES" THEN
	CALL navigate_to_MAXIS_screen("STAT", "MEMI")
	EmReadScreen marital_status, 1, 7, 40
	EmReadScreen spouse_ref_nbr, 02, 09, 49

	IF spouse_ref_nbr = "__" THEN
	    Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 176, 155, "Spouse, not found on MEMB"
  		EditBox 55, 5, 115, 15, spouse_first_name
  		EditBox 55, 25, 115, 15, spouse_last_name
  		EditBox 55, 45, 70, 15, spouse_SSN_number
  		EditBox 55, 65, 55, 15, spouse_DOB
  		DropListBox 55, 85, 55, 15, "Select One:"+chr(9)+"Female"+chr(9)+"Male"+chr(9)+"Unknown"+chr(9)+"Undetermined", spouse_gender_dropdown
  		EditBox 5, 115, 165, 15, other_notes
  		ButtonGroup ButtonPressed
    		OkButton 85, 135, 40, 15
    		CancelButton 130, 135, 40, 15
  		Text 5, 10, 40, 10, "First Name:"
  		Text 5, 50, 20, 10, "SSN: "
  		Text 5, 70, 45, 10, "Date of birth: "
  		Text 5, 105, 160, 10, "Please explain why they are not listed in maxis: "
  		Text 5, 90, 30, 10, "Gender: "
  		Text 5, 30, 40, 10, "Last Name: "
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

	ELSE
		CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	    EMwritescreen spouse_ref_nbr, 20, 76
		TRANSMIT
	    EMReadScreen spouse_first_name, 12, 6, 63
	    spouse_first_name = replace(spouse_first_name, "_", "")
	    spouse_first_name = trim(spouse_first_name)
	    EMReadScreen spouse_last_name, 25, 6, 30
	    spouse_last_name = replace(spouse_last_name, "_", "")
	    spouse_last_name = trim(spouse_last_name)
	    EMReadscreen spouse_SSN_number_read, 11, 7, 42
	    spouse_SSN_number_read = replace(spouse_SSN_number_read, " ", "")
	    EmReadScreen spouse_DOB, 10, 8, 42
	    spouse_DOB = replace(spouse_DOB, " ", "/")
	    EmReadScreen spouse_gender, 1, 9, 42
	END IF
END IF

CALL navigate_to_MAXIS_screen("STAT", "ADDR")
EMreadscreen resi_addr_line_one, 22, 6, 43
resi_addr_line_one = replace(resi_addr_line_one, "_", "")
EMreadscreen resi_addr_line_two, 22, 7, 43
resi_addr_line_two = replace(resi_addr_line_two, "_", "")
EMreadscreen resi_addr_city, 15, 8, 43
resi_addr_city = replace(resi_addr_city, "_", "")
EMreadscreen resi_addr_state, 2, 8, 66
EMreadscreen resi_addr_zip, 7, 9, 43
resi_addr_zip = replace(resi_addr_zip, "_", "")
EMreadscreen addr_county, 2, 9, 66

EMreadscreen mailing_addr_line_one, 22, 13, 43
mailing_addr_line_one = replace(mailing_addr_line_one, "_", "")
EMreadscreen mailing_addr_line_two, 22, 14, 43
mailing_addr_line_two = replace(mailing_addr_line_two, "_", "")
EMreadscreen mailing_addr_city, 15, 15, 43
mailing_addr_city = replace(mailing_addr_city, "_", "")
EmReadScreen mailing_addr_state, 2, 16, 43
mailing_addr_state = replace(mailing_addr_state, "_", "")
EMreadscreen mailing_addr_zip, 7, 16, 52
mailing_addr_zip = replace(mailing_addr_zip, "_", "")

'string for MAXIS address
maxis_addr = resi_addr_line_one & " " & resi_addr_line_two & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip
'string for mailing address
mail_MAXIS_addr = mailing_addr_line_one & " " & mailing_addr_line_two & " " & mailing_addr_city & " " & mailing_addr_state & " " & mailing_addr_zip

body_of_email = "A signed AVS form was received for-" & vbcr & "First name: " & client_first_name & vbcr & "Last name: " & client_last_name & vbcr & "AVS Form Received Date: " & avs_form_date & vbcr & "Social Security Number: " & client_SSN_number_read & vbcr & "Gender: " & client_gender & vbcr & "Date of birth: " & client_DOB & vbcr & "Application date: " & application_date & vbcr & "Address: " & resi_addr_line_one & resi_addr_line_two & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip & vbcr

If trim(mail_MAXIS_addr) <> "" then body_of_email = body_of_email & "Mailing address: " & mailing_addr_line_one & mailing_addr_line_two & " " & mailing_addr_city & " " & mailing_addr_state & " " & mailing_addr_zip & vbcr
body_of_email = body_of_email & "MA type: " & MA_type & vbcr & "HH size: " & HH_size & vbcr & "Applicant Type: " & applicant_type & vbcr & "Application Type: " & appl_type & vbcr
If spouse_deeming = "YES" then body_of_email = body_of_email & "Spouse: " & spouse_deeming & vbcr & "Spouse Member # " & spouse_ref_nbr & vbcr & "Spouse First Name: " & spouse_first_name & vbcr & "Spouse Last Name: " & spouse_last_name & vbcr & "Spouse Social Security Number: " & spouse_SSN_number_read & vbcr & "Spouse Gender: " & spouse_gender & vbcr & "Spouse Date of birth: " & spouse_DOB

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
CALL create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "AVS initial run requests  #" &  MAXIS_case_number & " Member # " & memb_number, body_of_email, "", TRUE)

script_end_procedure_with_error_report("The email has been created and sent to the QI Team.")
