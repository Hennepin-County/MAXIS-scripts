'GATHERING STATS===========================================================================================
name_of_script = "NOTICES - OUT OF STATE INQUIRY.vbs"
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("01/31/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
'Error proof functions
CALL check_for_MAXIS(TRUE)'if not in maxis fylo'
CALL MAXIS_case_number_finder(MAXIS_case_number)
CALL convert_date_into_MAXIS_footer_month(date, MAXIS_footer_month, MAXIS_footer_year)'can use this for any date MM/YY'

Dialog1 = ""
BEGINDIALOG Dialog1, 0, 0, 146, 105, "Out of State Inquiry"
 EditBox 55, 5, 55, 15, MAXIS_case_number
 DropListBox 55, 25, 85, 15, "Send", out_of_state_request '+chr(9)+"Received"+chr(9)+"Unknown/No Response"'
 DropListBox 55, 45, 85, 15, "Select One:"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Phone"+chr(9)+"PARIS", how_sent
 DropListBox 55, 65, 85, 15, "Select One:"+chr(9)+"Alabama"+chr(9)+"Alaska"+chr(9)+"Arizona"+chr(9)+"Arkansas"+chr(9)+"California"+chr(9)+"Colorado"+chr(9)+"Connecticut"+chr(9)+"Delaware"+chr(9)+"Florida"+chr(9)+"Georgia"+chr(9)+"Hawaii"+chr(9)+"Idaho"+chr(9)+"Illinois"+chr(9)+"Indiana"+chr(9)+"Iowa"+chr(9)+"Kansas"+chr(9)+"Kentucky"+chr(9)+"Louisiana"+chr(9)+"Maine"+chr(9)+"Maryland"+chr(9)+"Massachusetts"+chr(9)+"Michigan"+chr(9)+"Mississippi"+chr(9)+"Missouri"+chr(9)+"Montana"+chr(9)+"Nebraska"+chr(9)+"Nevada"+chr(9)+"New Hampshire"+chr(9)+"New Jersey"+chr(9)+"New Mexico"+chr(9)+"New York"+chr(9)+"North Carolina"+chr(9)+"North Dakota"+chr(9)+"Ohio"+chr(9)+"Oklahoma"+chr(9)+"Oregon"+chr(9)+"Pennsylvania"+chr(9)+"Rhode Island"+chr(9)+"South Carolina"+chr(9)+"South Dakota"+chr(9)+"Tennessee"+chr(9)+"Texas"+chr(9)+"Utah"+chr(9)+"Vermont"+chr(9)+"Virginia"+chr(9)+"Washington"+chr(9)+"West Virginia"+chr(9)+"Wisconsin"+chr(9)+"Wyoming", state_droplist
 ButtonGroup ButtonPressed
   OkButton 55, 85, 40, 15
   CancelButton 100, 85, 40, 15
   Text 5, 10, 50, 10, "Case Number:"
   Text 20, 30, 30, 10, "Request:"
   Text 20, 50, 30, 10, "Via(How):"
   Text 30, 70, 20, 10, "State:"
ENDDIALOG

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "Enter a valid case number."
		If state_droplist = "Select One:" then err_msg = err_msg & vbnewline & "Select the state."
		If how_sent = "Select One:" then err_msg = err_msg & vbnewline & "Select how the request was sent."
		If out_of_state_request = "Select One:" then err_msg = err_msg & vbNewLine & "Please select the status of the out of state inquiry."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)                                 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false

CALL navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) 'gives a true/false'
IF is_this_priv = TRUE THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
'new function set to true to navigate and let the worker know if they need to request access '
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
row = 1                                                 'First we will look for SNAP
col = 1
EMSearch "IV-E:", row, col
If row <> 0 Then
	EMReadScreen IVE_status, 9, row, col + 6
	IVE_status = trim(IVE_status)
 	IF IVE_status = "ACTIVE" or IVE_status = "APP CLOSE" or IVE_status = "APP OPEN" Then
 		ive_case = TRUE
		case_active = TRUE
	END IF
	If IVE_status = "PENDING" Then
		ive_case = TRUE
		case_pending = TRUE
 	END IF
END IF
row = 1                                                 'First we will look for SNAP
col = 1
EMSearch "CCAP:", row, col
IF row <> 0 Then
  	EMReadScreen CC_status, 9, row, col + 6
  	CC_status = trim(CC_status)
  	IF CC_status = "ACTIVE" or CC_status = "APP CLOSE" or CC_status = "APP OPEN" Then
  		ccap_case = TRUE
  		case_active = TRUE
    END IF
  	IF CC_status = "PENDING" Then
  	  	ccap_case = TRUE
  		case_pending = TRUE
  	END IF
END IF 'reminder to look at adding thse to function at a later date'

IF case_active = FALSE and case_pending = FALSE THEN script_end_procedure_with_error_report("It appears no programs are open or pending on this case.")

IF family_cash_case = TRUE THEN MN_CASH_CHECKBOX = CHECKED
IF adult_cash_case = TRUE THEN MN_CASH_CHECKBOX = CHECKED
IF unknown_cash_pending = TRUE THEN MN_CASH_CHECKBOX = CHECKED
IF grh_case = TRUE THEN MN_GRH_CHECKBOX = CHECKED
IF snap_case = TRUE THEN MN_FS_CHECKBOX = CHECKED
IF ma_case = TRUE THEN  MN_HC_CHECKBOX = CHECKED
IF msp_case = TRUE THEN  MN_HC_CHECKBOX = CHECKED
IF ccap_case = TRUE THEN MN_CCA_CHECKBOX = CHECKED
'IF ive_case = TRUE THEN MN_IVE_CHECKBOX = CHECKED need to add to dialog


CALL navigate_to_MAXIS_screen("STAT", "ADDR")
EMReadScreen client_1staddress, 21, 06, 43
EMReadScreen client_2ndaddress, 21, 07, 43
EMReadScreen client_city, 14, 08, 43
EMReadScreen client_state, 2, 08, 66
EMReadScreen client_zip, 7, 09, 43
EMreadscreen addr_homeless, 1, 10, 43
client_address = replace(client_1staddress, "_","") & " " & replace(client_2ndaddress, "_","") & " " & replace(client_city, "_","") & ", " & replace(client_state, "_","") & " " & replace(client_zip, "_","")
EMreadscreen addr_homeless, 1, 10, 43
Const ref_numb_const 			= 0
Const first_name_const 			= 1
Const last_name_const			= 2
Const clt_middle_const 			= 3
Const clt_dob_const 			= 4
Const client_selection_checkbox_const = 5
Const clt_ssn_const 			= 6

Dim ALL_CLT_INFO_ARRAY()
ReDim ALL_CLT_INFO_ARRAY(clt_ssn_const, 0)

the_incrementer = 0
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2
	ReDim Preserve ALL_CLT_INFO_ARRAY(clt_ssn_const, the_incrementer)' this is to tell the array to get bigger'
	ALL_CLT_INFO_ARRAY(ref_numb_const, the_incrementer) = ref_nbr 'going back to the first piece of information to hold it in this specific postion'
	'MsgBox access_denied_check
	If access_denied_check = "ACCESS DENIED" Then
		PF10
		last_name = "UNABLE TO FIND"
		first_name = " - Access Denied"
		mid_initial = ""
	Else
	    'Reading info and removing spaces
	    EMReadscreen First_name, 12, 6, 63
	    First_name = replace(First_name, "_", "")
	    ALL_CLT_INFO_ARRAY(first_name_const, the_incrementer) = First_name
	    'Reading Last name and removing spaces
	    EMReadscreen Last_name, 25, 6, 30
	    Last_name = replace(Last_name, "_", "")
	    ALL_CLT_INFO_ARRAY(last_name_const, the_incrementer) = Last_name
	    'Reading Middle initial and replacing _ with a blank if empty.
	    EMReadscreen Middle_initial, 1, 6, 79
	    Middle_initial = replace(Middle_initial, "_", "")
	    ALL_CLT_INFO_ARRAY(clt_middle_const, the_incrementer) = Middle_initial
		 'Reading date of birth and replacing space.
	    Emreadscreen client_dob, 10, 8, 42
	    SSN_number = replace(client_dob, " ", "/")
	    ALL_CLT_INFO_ARRAY(clt_dob_const, the_incrementer) = client_dob
	    'Reads SSN
	    Emreadscreen SSN_number, 11, 7, 42
	    SSN_number = replace(SSN_number, " ", "-")
	    ALL_CLT_INFO_ARRAY(clt_ssn_const, the_incrementer) = SSN_number
		'adds the ref number to the array'
	    ALL_CLT_INFO_ARRAY(ref_numb_const, the_incrementer) = client_ref_number
		'ensuring that the check box is checked for all members in the dialog'
		ALL_CLT_INFO_ARRAY(client_selection_checkbox_const, the_incrementer) = CHECKED
	End If
	the_incrementer = the_incrementer + 1
	TRANSMIT
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

'For each path the script takes a different route'
Dialog1 = "" 'runs the dialog that has been dynamically created. Streamlined with new functions.
BEGINDIALOG Dialog1, 0, 0, 241, (50 + (Ubound(ALL_CLT_INFO_ARRAY, 2) * 15)), "Household Member(s) "   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 130, 10, "Select household members to request:"
	FOR the_pers = 0 to Ubound(ALL_CLT_INFO_ARRAY, 2)
		checkbox 10, (20 + (the_pers * 15)), 160, 10, ALL_CLT_INFO_ARRAY(ref_numb_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(first_name_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(last_name_const, the_pers) & " " & ALL_CLT_INFO_ARRAY(clt_ssn_const, the_pers), ALL_CLT_INFO_ARRAY(client_selection_checkbox_const, the_pers)
	NEXT
	ButtonGroup ButtonPressed
	OkButton 185, 10, 50, 15
	CancelButton 185, 30, 50, 15
ENDDIALOG

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
	LOOP until err_msg = ""
Loop until are_we_passworded_out = false

Call fill_in_the_states
'TODO go through the functions and agency_email = agency_email ; &

date_received = ""
