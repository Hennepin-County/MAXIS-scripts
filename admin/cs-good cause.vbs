'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CS GOOD CAUSE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 240          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("06/13/2018", "Updated incomplete forms and dialog.", "MiKayla Handley, Hennepin County")
call changelog_update("05/14/2018", "Updated per GC Committee requests.", "MiKayla Handley, Hennepin County")
call changelog_update("03/27/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Fun with dates! --Creating variables for the rolling 12 calendar months
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month -2'
CM_minus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
CM_minus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)
'current month -3'
CM_minus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", -3, date)            ), 2)
CM_minus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", -3, date)            ), 2)
'current month -4'
CM_minus_4_mo =  right("0" &             DatePart("m",           DateAdd("m", -4, date)            ), 2)
CM_minus_4_yr =  right(                  DatePart("yyyy",        DateAdd("m", -4, date)            ), 2)
'current month -5'
CM_minus_5_mo =  right("0" &             DatePart("m",           DateAdd("m", -5, date)            ), 2)
CM_minus_5_yr =  right(                  DatePart("yyyy",        DateAdd("m", -5, date)            ), 2)
'current month -6'
CM_minus_6_mo =  right("0" &             DatePart("m",           DateAdd("m", -6, date)            ), 2)
CM_minus_6_yr =  right(                  DatePart("yyyy",        DateAdd("m", -6, date)            ), 2)
'current month -7'
CM_minus_7_mo =  right("0" &             DatePart("m",           DateAdd("m", -7, date)            ), 2)
CM_minus_7_yr =  right(                  DatePart("yyyy",        DateAdd("m", -7, date)            ), 2)
'current month -8'
CM_minus_8_mo =  right("0" &             DatePart("m",           DateAdd("m", -8, date)            ), 2)
CM_minus_8_yr =  right(                  DatePart("yyyy",        DateAdd("m", -8, date)            ), 2)
'current month -9'
CM_minus_9_mo =  right("0" &             DatePart("m",           DateAdd("m", -9, date)            ), 2)
CM_minus_9_yr =  right(                  DatePart("yyyy",        DateAdd("m", -9, date)            ), 2)
'current month -10'
CM_minus_10_mo =  right("0" &            DatePart("m",           DateAdd("m", -10, date)           ), 2)
CM_minus_10_yr =  right(                 DatePart("yyyy",        DateAdd("m", -10, date)           ), 2)
'current month -11'
CM_minus_11_mo =  right("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'Establishing value of variables for the rolling 12 months
current_month = CM_mo & "/" & CM_yr
current_month_minus_one = CM_minus_1_mo & "/" & CM_minus_1_yr
current_month_minus_two = CM_minus_2_mo & "/" & CM_minus_2_yr
current_month_minus_three = CM_minus_3_mo & "/" & CM_minus_3_yr
current_month_minus_four = CM_minus_4_mo & "/" & CM_minus_4_yr
current_month_minus_five = CM_minus_5_mo & "/" & CM_minus_5_yr
current_month_minus_six = CM_minus_6_mo & "/" & CM_minus_6_yr
current_month_minus_seven = CM_minus_7_mo & "/" & CM_minus_7_yr
current_month_minus_eight = CM_minus_8_mo & "/" & CM_minus_8_yr
current_month_minus_nine = CM_minus_9_mo & "/" & CM_minus_9_yr
current_month_minus_ten = CM_minus_10_mo & "/" & CM_minus_10_yr
current_month_minus_eleven = CM_minus_11_mo & "/" & CM_minus_11_yr

'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
'Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
MEMB_number = "01"

'----------------------------------------------------------------------------------------------------DIALOGS
BeginDialog change_exemption_dialog, 0, 0, 216, 100, "Good cause change/exemption "
  EditBox 110, 5, 50, 15, change_reported_date
  EditBox 110, 25, 100, 15, change_reported
  EditBox 110, 45, 100, 15, maxis_updates
  CheckBox 10, 70, 75, 10, " No longer claiming ", no_longer_claiming_checkbox
  ButtonGroup ButtonPressed
    OkButton 110, 80, 50, 15
    CancelButton 160, 80, 50, 15
  Text 10, 10, 80, 10, "Date of change reported:"
  Text 10, 30, 90, 10, "What change was reported:"
  Text 10, 50, 95, 10, "What was updated in MAXIS:"
EndDialog

BeginDialog good_cause_dialog, 0, 0, 386, 285, "Good Cause"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 165, 5, 40, 15, actual_date
  EditBox 80, 25, 20, 15, memb_number
  EditBox 165, 25, 40, 15, claim_date
  EditBox 80, 45, 20, 15, child_memb_number
  EditBox 165, 45, 40, 15, review_date
  CheckBox 225, 10, 70, 10, "Med Sup Svc Only", med_sup_check
  CheckBox 225, 25, 60, 10, "Sup Evidence", sup_evidence_check
  CheckBox 225, 40, 55, 10, "Investigation", investigation_check
  CheckBox 320, 15, 25, 10, "CCA", CCA_CHECKBOX
  CheckBox 320, 30, 30, 10, "DWP", DWP_CHECKBOX
  CheckBox 320, 45, 30, 10, "METS", METS_CHECKBOX
  CheckBox 355, 15, 20, 10, "FS", FS_CHECKBOX
  CheckBox 355, 30, 20, 10, "HC", HC_CHECKBOX
  CheckBox 355, 45, 20, 10, "MF", MFIP_CHECKBOX
  CheckBox 10, 95, 145, 10, "ABPS name not written on the correct line", ABPS_CHECKBOX
  CheckBox 10, 105, 140, 10, "Reason for requesting GC not selected", REASON_CHECKBOX
  CheckBox 165, 95, 120, 10, "All of the questions not answered", QUESTIONS_CHECKBOX
  CheckBox 165, 105, 90, 10, "No signature and/or date ", NOSIG_CHECKBOX
  CheckBox 290, 105, 80, 10, "Other (please specify)", OTHER_CHECKBOX
  DropListBox 30, 65, 55, 15, "Select One:"+chr(9)+"Denied"+chr(9)+"Granted"+chr(9)+"Not Claimed"+chr(9)+"Pending", gc_status
  DropListBox 115, 65, 115, 15, "Select One:"+chr(9)+"Application Review-Complete"+chr(9)+"Application Review-Incomplete"+chr(9)+"Change/exemption ending"+chr(9)+"Determination"+chr(9)+"Recertification", good_cause_droplist
  DropListBox 265, 65, 115, 15, "Select One:"+chr(9)+"Potential phys harm/Child"+chr(9)+"Potential Emotnl harm/Child"+chr(9)+"Potential phys harm/Caregiver"+chr(9)+"Potential Emotnl harm/Caregiver"+chr(9)+"Cncptn Incest/Forced Rape"+chr(9)+"Legal adoption Before Court"+chr(9)+"Parent Gets Preadoptn Svc"+chr(9)+"No Longer Claiming", reason_droplist
  EditBox 65, 125, 85, 15, mets_info
  EditBox 65, 145, 85, 15, verifs_req
  EditBox 210, 125, 170, 15, denial_reason
  EditBox 210, 145, 170, 15, other_notes
  CheckBox 10, 175, 185, 10, "Sent Request for Proof to Support Good Cause Claim", SUP_CHECKBOX
  CheckBox 10, 185, 160, 10, "Sent Good Cause Client Statement (DHS-2338)", DHS_2338_CHECKBOX
  CheckBox 10, 195, 220, 10, "Sent Imp Information about Your Request Exemption (DHS-3627) ", DHS_3627_CHECKBOX
  CheckBox 10, 205, 205, 10, "Sent Notice of Denial of Good Cause Exemption (DHS-3628) ", DHS_3628_CHECKBOX
  CheckBox 10, 215, 165, 10, "Sent Notice of Good Cause Approval (DHS-3629) ", DHS_3629_CHECKBOX
  CheckBox 10, 225, 195, 10, "Sent Request to End Good Cause Exemption  (DHS-3631 )", DHS_3631_CHECKBOX
  CheckBox 10, 235, 180, 10, "Sent Request for Additional Information (DHS 3632)", DHS_3632_CHECKBOX
  CheckBox 10, 245, 165, 10, "Sent Good Cause Yearly Determination Packet", Recert_CHECKBOX
  CheckBox 10, 265, 240, 10, "Good Cause Client Statement (DHS-2338) is in ECF and completed in full", DHS_2338_complete_CHECKBOX
  CheckBox 10, 255, 195, 10, "Sent Good Cause Redetermination Approval ( DHS 3633)", DHS_3633_CHECKBOX
  ButtonGroup ButtonPressed
    OkButton 270, 265, 50, 15
    CancelButton 325, 265, 50, 15
  Text 120, 10, 40, 10, "Actual date:"
  Text 90, 70, 25, 10, "Action:"
  Text 235, 70, 30, 10, "Reason:"
  Text 5, 70, 25, 10, "Status:"
  Text 5, 130, 60, 10, "Mets Information:"
  Text 5, 150, 60, 10, "Verifs Requested:"
  Text 155, 130, 50, 10, "Denial Reason:"
  Text 165, 150, 45, 10, "Other Notes:"
  Text 115, 50, 45, 10, "Next review:"
  GroupBox 315, 5, 65, 55, "Programs"
  GroupBox 5, 85, 375, 35, "Incomplete Form:"
  Text 5, 10, 45, 10, "Case number:"
  Text 120, 30, 40, 10, "Claim date:"
  Text 10, 50, 65, 10, "Child's MEMB #(s):"
  GroupBox 5, 165, 250, 115, "Verifications"
  Text 5, 30, 75, 10, "Caregiver's MEMB #:"
EndDialog

'Initial dialog giving the user the option to select the type of good cause action
MAXIS_case_number = "276348"
actual_date = "09/01/18"
memb_number = "01"
claim_date = "09/15/18"
child_memb_number = "03, 04"
review_date = "09/15/19"

Do
	Do
		err_msg = ""
		dialog good_cause_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF memb_number = "" THEN err_msg = err_msg & vbnewline & "* Please enter the caregiver member number."
		IF child_memb_number = "" THEN err_msg = err_msg & vbnewline & "* Please enter the child's member number."
		If good_cause_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the Good Cause Application Status."
		IF gc_status = "Select One:" THEN err_msg = err_msg & vbnewline & "* Select the Good Cause Case Status."
		If reason_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the Good Cause reason."
		If gc_status = "Granted" THEN
			If isdate(review_date) = False then err_msg = err_msg & vbnewline & "* You must enter a valid good cause review date."
		ELSEIF gc_status = "Denied" THEN
			If denial_reason = "" then err_msg = err_msg & vbnewline & "* You must enter a denial reason."
		END IF
		If isdate(actual_date) = FALSE then err_msg = err_msg & vbnewline & "* You must enter an actual date in the footer month that you are working in."
		IF gc_status <> "Not Claimed" THEN
			If isdate(claim_date) = False then err_msg = err_msg & vbnewline & "* You must enter a valid good cause claim date."
		END IF
		If good_cause_droplist = "Application Review-Incomplete" then
			If other_notes = "" and OTHER_CHECKBOX = CHECKED then err_msg = err_msg & vbnewline & "* You must enter a reason that application is incomplete."
		END IF
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If good_cause_droplist = "Change/exemption ending" THEN
	Do
		Do
			err_msg = ""
			dialog change_exemption_dialog
			cancel_confirmation
			If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
MsgBox actual_date
the_month = datepart("m", actual_date)
MAXIS_footer_month = right("00" & the_month, 2)
the_year = datepart("yyyy", actual_date)
MAXIS_footer_year = right("00" & the_year, 2)
CALL convert_date_into_MAXIS_footer_month(actual_date, footer_month, footer_year)
	'----------------------------------------------------------------------------------------------------ABPS panel
Call navigate_to_MAXIS_screen("STAT", "ABPS")
	'Making sure we have the correct ABPS
DO
	EMReadScreen panel_number, 1, 2, 78
	EMReadScreen abps_number, 9, 13, 40
	abps_number = trim(abps_number)
	If panel_number = "0" then script_end_procedure("An ABPS panel does not exist. Please create the panel before running the script again. ")
	'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
	Do
		EMReadScreen current_panel_number, 1, 2, 73
		ABPS_check = MsgBox("Is this the right ABPS " & abps_number & "?", vbYesNo + vbQuestion, "Confirmation of ABPS")
		If ABPS_check = vbYes then exit do
		If ABPS_check = vbNo then TRANSMIT
		If (ABPS_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure("Unable to find another ABPS. Please review the case, and run the script again if applicable.")
	Loop until current_panel_number = panel_number
	'-------------------------------------------------------------------------Updating the ABPS panel


	Do
		EMReadScreen first_display_mode_check, 1, 20, 8
		IF first_display_mode_check = "E" then exit do
		IF first_display_mode_check = "D" THEN
			PF9'edit mode
			EMReadScreen error_msg, 2, 24, 2	'making sure we can actually update this case.
			error_msg = trim(error_msg)
			If error_msg <> "" then script_end_procedure("Unable to update this case. Please review case, and run the script again if applicable.")
		END IF
	Loop until first_display_mode_check = "E"
	Call create_MAXIS_friendly_date_with_YYYY(actual_date, 0, 6, 35)
	EMWriteScreen "Y", 4, 73			'Support Coop Y/N field
	msgbox "I entered a Y"
	IF gc_status = "Pending" THEN
		Msgbox gc_status
		EMWriteScreen "P", 5, 47			'Good Cause status field
		EMWriteScreen "  ", 6, 73'next review date'
		EMWriteScreen "  ", 6, 76'next review date'
		EMWriteScreen "  ", 6, 79'next review date'
	ELSEIF gc_status = "Granted" THEN
		Msgbox gc_status
		EMWriteScreen "G", 5, 47
		Call create_MAXIS_friendly_date(datevalue(review_date), 0, 6, 73)
	ELSEIF gc_status = "Denied" THEN
		Msgbox gc_status
		EMWriteScreen "D", 5, 47
		EMWriteScreen "  ", 6, 73'next review date'
		EMWriteScreen "  ", 6, 76'next review date'
		EMWriteScreen "  ", 6, 79'next review date'
	ELSEIF gc_status = "Not Claimed" THEN
		Msgbox gc_status
		EMWriteScreen "N", 5, 47
		EMWriteScreen "  ", 5, 73'good cause claim date'
		EMWriteScreen "  ", 5, 76'good cause claim date'
		EMWriteScreen "  ", 5, 79'good cause claim date'
		EMWriteScreen " ", 6, 47'reason good cause claimed'
		EMWriteScreen "  ", 6, 73'next review date'
		EMWriteScreen "  ", 6, 76'next review date'
		EMWriteScreen "  ", 6, 79'next review date'
		EMWriteScreen " ", 7, 47'Sup Evidence'
		EMWriteScreen " ", 7, 73 'Investigation'
		EMWriteScreen " ", 8, 48 'Med Sup Svc Only'
	END IF
	'converting the good cause reason from reason_droplist to the applicable MAXIS coding
	If reason_droplist = "Potential phys harm/Child"		then claim_reason = "1"
	If reason_droplist = "Potential Emotnl harm/Child"	 	then claim_reason = "2"
	If reason_droplist = "Potential phys harm/Caregiver" 	then claim_reason = "3"
	If reason_droplist = "Potential Emotnl harm/Caregiver" 	then claim_reason = "4"
	If reason_droplist = "Cncptn Incest/Forced Rape" 		then claim_reason = "5"
	If reason_droplist = "Legal adoption Before Court" 		then claim_reason = "6"
	If reason_droplist = "Parent Gets Preadoptn Svc" 		then claim_reason = "7"
	IF gc_status <> "Not Claimed" THEN
		Call create_MAXIS_friendly_date(datevalue(claim_date), 0, 5, 73)
		IF sup_evidence_CHECKBOX = CHECKED THEN EMWriteScreen "Y", 7, 47 ELSE EMWriteScreen "N", 7, 47
		IF investigation_CHECKBOX = CHECKED THEN EMWriteScreen "Y", 7, 73 ELSE EMWriteScreen "N", 7, 73
		IF med_sup_CHECKBOX = CHECKED THEN EMWriteScreen "Y", 8, 48 ELSE EMWriteScreen "N", 8, 48
		EMWriteScreen claim_reason, 6, 47
	END IF
	Call create_MAXIS_friendly_date_with_YYYY(datevalue(actual_date), 0, 18, 38) 'creates and writes the date entered in dialog'
	EMReadScreen parental_status, 1, 15, 53	'making sure ABPS is not unknown.
	msgbox parental_status
	IF parental_status = "2" THEN
		client_name = "Unknown"
	ELSEIF parental_status = "3" THEN
		client_name = "ABPS deceased"
	ELSEIF parental_status = "4" THEN
		client_name = "Rights Severed"
	ELSEIF parental_status = "7" THEN
		client_name = "HC No Order Sup"
	ELSEIF parental_status = "1" THEN
		EMReadScreen first_name, 12, 10, 63
		EMReadScreen last_name, 24, 10, 30
		first_name = trim(first_name)
		last_name = trim(last_name)
		first_name = replace(first_name, "_", "")
		last_name = replace(last_name, "_", "")
		client_name = first_name & " " & last_name
		Call fix_case_for_name(client_name)
	END IF

	Transmit'to add information
	Transmit'to move past non-inhibiting warning messages on ABPS (this also takes you to the next abps)
	PF3
	EMReadScreen ABPS_screen, 4, 2, 50		'if inhibiting error exists, this will catch it and instruct the user to update ABPS
	msgbox ABPS_screen
	If ABPS_screen = "ABPS" then script_end_procedure("An error occurred on the ABPS panel. Please update the panel before using the script with the absent parent information.")

	EMReadScreen error_msg, 7, 24, 2	'making sure we can actually update this case.
	error_msg = trim(error_msg)
	If error_msg = "WARNING" then TRANSMIT
	'--------------------------------------------------------------------------------this will run the case thru background 'and update the future months
	EMReadScreen WRAP_check, 4, 2, 46
	If WRAP_check = "WRAP" then
		EMReadScreen MAXIS_footer_month, 2, 20, 55
		EMReadScreen MAXIS_footer_year, 2, 20, 58
		If datediff("m", date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 1 then in_future_month = True
		If future_months_check = 0 or in_future_month = True then exit do
		IF in_future_month = FALSE THEN
		 	CALL write_value_and_transmit("y", 16, 54)
			EMWriteScreen "ABPS", 20, 71
			If len(current_panel_number) = 1 then current_panel_number = "0" & current_panel_number
			EMWriteScreen current_panel_number, 20, 79
			transmit
			PF9
		END IF
	END IF
Loop until in_future_month = True

'Case note stuffs'
IF CCA_CHECKBOX = CHECKED THEN programs_included = programs_included & "CCAP, "
IF DWP_CHECKBOX = CHECKED THEN programs_included = programs_included & "DWP, "
IF HC_CHECKBOX = CHECKED THEN programs_included = programs_included & "Healthcare, "
IF FS_CHECKBOX = CHECKED THEN programs_included = programs_included & "Food Support, "
IF MFIP_CHECKBOX = CHECKED THEN programs_included = programs_included & "MFIP, "
IF METS_CHECKBOX = CHECKED THEN programs_included = programs_included & "MNSURE, "
'trims excess spaces of programs
programs_included  = trim(programs_included )
'takes the last comma off of programs
If right(programs_included, 1) = "," THEN programs_included  = left(programs_included, len(programs_included) - 1)
IF ABPS_CHECKBOX = CHECKED THEN incomplete_form = incomplete_form & "ABPS name not written on the correct line,"
IF REASON_CHECKBOX = CHECKED THEN incomplete_form = incomplete_form & " reason for requesting GC not selected,"
IF QUESTIONS_CHECKBOX = CHECKED THEN incomplete_form = incomplete_form & " all of the questions not answered,"
IF NOSIG_CHECKBOX = CHECKED THEN incomplete_form = incomplete_form & " no signature and/or date,"
IF OTHER_CHECKBOX = CHECKED THEN incomplete_form = incomplete_form & " other (see additional information),"
incomplete_form  = trim(incomplete_form)
If right(incomplete_form, 1) = "," THEN incomplete_form  = left(incomplete_form, len(incomplete_form) - 1)

'-----------------------------------------------------------------------------------------------------Case note & email sending
start_a_blank_CASE_NOTE
IF good_cause_droplist = "Application Review-Complete" THEN Call write_variable_in_case_note("Good Cause Application Review - Complete")
IF good_cause_droplist = "Application Review-Incomplete" THEN Call write_variable_in_case_note("Good Cause Application Review - Incomplete")
IF good_cause_droplist = "Change/exemption ending" THEN
	Call write_variable_in_case_note("Good Cause Application Change/exemption ending")
	Call write_bullet_and_variable_in_case_note("Date of change reported", change_reported_date)
	Call write_bullet_and_variable_in_case_note("What change was reported", change_reported)
	Call write_bullet_and_variable_in_case_note("What was updated in MAXIS", maxis_updates)
	IF no_longer_claiming_checkbox = CHECKED THEN Call write_variable_in_case_note("* Client is no longer claiming good cause")
END IF
IF good_cause_droplist = "Determination" THEN Call write_variable_in_case_note("Good Cause Application - Determination")
IF good_cause_droplist = "Recertification" THEN Call write_variable_in_case_note("Good Cause Application Review - Recertification")
Call write_bullet_and_variable_in_case_note("Good cause status", gc_status)
If claim_date <> "" THEN Call write_bullet_and_variable_in_case_note("Good cause claim date", claim_date)
If review_date <> "" THEN Call write_bullet_and_variable_in_case_note("Next review date", review_date)
Call write_bullet_and_variable_in_case_note("Child(ren) member number(s)", child_memb_number)
Call write_bullet_and_variable_in_case_note("ABPS name", client_name)
CALL write_bullet_and_variable_in_case_note("Applicable programs", programs_included)
 	IF reason_droplist <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Reason for claiming good cause", reason_droplist)
IF incomplete_form <> "Select One" THEN Call write_bullet_and_variable_in_case_note("What is GC form incomplete for", incomplete_form)
If denial_reason <> "" THEN Call write_bullet_and_variable_in_case_note("Reason for denial", denial_reason)
IF mets_info <> "" THEN Call write_bullet_and_variable_in_case_note("METS information", mets_info )
IF verfis_req <> "" THEN Call write_bullet_and_variable_in_case_note("Requested Verifcation(s)", verifs_req)
IF other_notes <> "" THEN Call write_bullet_and_variable_in_case_note("Additional information", other_notes)
IF DHS_2338_complete_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* DHS-2338 is in ECF, and fully completed by parent/caregiver.")
IF SUP_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent request of proof to support a good cause claim")
IF DHS_2338_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Good Cause Client Statement (DHS-2338)")
IF DHS_3628_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Notice of Denial of Good Cause Exemption (DHS-3628)")
IF DHS_3629_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Notice of Good Cause Approval (DHS-3629)")
IF DHS_3632_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Request for Additional Information (DHS 3632)")
IF DHS_3631_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Good Cause End Exemption (DHS-3631)")
IF DHS_3627_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Imp Information about Your Request Exemption (DHS-3627)")
IF Recert_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Good Cause Yearly Determination Packet")
IF DHS_3633_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Good Cause Redetermination Approval (DHS 3633)")
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)
IF FS_CHECKBOX = CHECKED and CCA_CHECKBOX = UNCHECKED and DWP_CHECKBOX = UNCHECKED and MFIP_CHECKBOX = UNCHECKED and HC_CHECKBOX = UNCHECKED and METS_CHECKBOX = UNCHECKED THEN memo_started = TRUE
IF memo_started = TRUE THEN
	Call start_a_new_spec_memo
	EMsendkey("************************************************************")
	Call write_variable_in_SPEC_MEMO("You recently applied for Food Support assistance and")
	Call write_variable_in_SPEC_MEMO("requested Good Cause for Child Support.")
	Call write_variable_in_SPEC_MEMO("You do not need to cooperate with Child Support for Food")
	Call write_variable_in_SPEC_MEMO("Support applications, therefore you do not need to request")
	Call write_variable_in_SPEC_MEMO("good cause at this time.")
	Call write_variable_in_SPEC_MEMO("If you apply for Cash or Health Care programs in the future")
	Call write_variable_in_SPEC_MEMO("you will need to resubmit the application for Good Cause.")
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4
END IF

script_end_procedure("Success! MAXIS has been updated, and the Good Cause results case noted.")
