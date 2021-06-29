'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - CS GOOD CAUSE.vbs"
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
call changelog_update("09/09/2019", "Handling for edit mode on ABPS panel.", "MiKayla Handley, Hennepin County")
call changelog_update("06/13/2018", "Updated incomplete forms and dialog.", "MiKayla Handley, Hennepin County")
call changelog_update("05/14/2018", "Updated per GC Committee requests.", "MiKayla Handley, Hennepin County")
call changelog_update("03/27/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""
'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = datepart("m", date)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("yyyy", date), 2)

'Grabbing the case number
call MAXIS_case_number_finder(MAXIS_case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
back_to_self 'to ensure we are not in edit mode'
EMReadScreen priv_check, 4, 24, 18 'If it can't get into the case needs to skip
IF priv_check = "INVA" THEN script_end_procedure_with_error_report("This case is invalid for period selected. ")

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 116, 65, "Case number"
  EditBox 60, 5, 50, 15, MAXIS_case_number
  EditBox 65, 25, 20, 15, MAXIS_footer_month
  EditBox 90, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 5, 45, 50, 15
    CancelButton 60, 45, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 45, 10, "Footer Month:"
EndDialog
DO
	DO
	    err_msg = ""
	    Dialog Dialog1
	    cancel_confirmation
	    IF MAXIS_case_number = "" THEN err_msg = "Please enter a case number to continue."
	    IF MAXIS_footer_month = "" THEN err_msg = "Please enter a footer month to continue."
	    IF MAXIS_footer_year = "" THEN err_msg = "Please enter a footer year to continue."
	    IF err_msg <> "" THEN msgbox "*** Error Check ***" & vbNewLine & err_msg
		LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

EMWriteScreen MAXIS_case_number, 18, 43
Call navigate_to_MAXIS_screen("CASE", "CURR")
EMReadScreen CURR_panel_check, 4, 2, 55
EMReadScreen case_status, 8, 8, 9
case_status = trim(case_status)

IF case_status = "ACTIVE" THEN active_status = TRUE
IF case_status = "APP OPEN" THEN active_status = TRUE
IF case_status = "APP CLOS" THEN active_status = TRUE
IF case_status = "INACTIVE" THEN active_status = FALSE
If case_status = "CAF2 PEN" THEN active_status = TRUE
If case_status = "CAF1 PEN" THEN active_status = TRUE
IF case_status = "REIN" THEN active_status = TRUE

Call MAXIS_footer_month_confirmation
EmReadscreen original_MAXIS_footer_month, 2, 20, 43
'msgbox original_MAXIS_footer_month
CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'Checking for PRIV cases.
EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
IF priv_check = "PRIV" THEN script_end_procedure_with_error_report("This case is privileged. Please request access before running the script again. ")
CASH_STATUS = FALSE 'overall variable'
CCA_STATUS = FALSE
DWP_STATUS = FALSE 'Diversionary Work Program'
ER_STATUS = FALSE
FS_STATUS = FALSE
GA_STATUS = FALSE 'General Assistance'
GRH_STATUS = FALSE
HC_STATUS = FALSE
MS_STATUS = FALSE 'Mn Suppl Aid '
MF_STATUS = FALSE 'Mn Family Invest Program '
RC_STATUS = FALSE 'Refugee Cash Assistance'

'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
'MsgBox  cash1_status_check
EMReadScreen cash2_status_check, 4, 7, 74
'MsgBox cash2_status_check
EMReadScreen emer_status_check, 4, 8, 74
'MsgBox emer_status_check
EMReadScreen grh_status_check, 4, 9, 74
'MsgBox grh_status_check
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
'MsgBox cca_status_check
EMReadScreen cash1_prog_check, 2, 6, 67
'MsgBox cash1_prog_check
EMReadScreen cash2_prog_check, 2, 7, 67
'MsgBox cash2_prog_check
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND" THEN
	FS_STATUS = TRUE
	FS_CHECKBOX = CHECKED
END IF

IF hc_status_check = "ACTV" or hc_status_check = "PEND" THEN
	HC_STATUS = TRUE
	HC_CHECKBOX   = CHECKED
END IF

IF cca_status_check = "ACTV" or cca_status_check = "PEND" THEN
	CCA_STATUS = TRUE
	CCA_CHECKBOX  = CHECKED
END IF
'Logic to determine if MFIP is active
IF cash1_prog_check = "MF" THEN
	IF cash1_status_check = "ACTV" or cash1_status_check = "PEND" THEN
		MF_STATUS = TRUE
		'MsgBox MF_STATUS
		MFIP_CHECKBOX = CHECKED
		'MsgBox MFIP_CHECKBOX
	END IF
	If cash1_status_check = "INAC" or cash1_status_check = "SUSP" or cash1_status_check = "DENY" or cash1_status_check = "" THEN MF_STATUS = FALSE
END IF

IF cash1_prog_check = "MF" THEN
	IF cash2_status_check = "ACTV" or cash2_status_check = "PEND" THEN
		MF_STATUS = TRUE
		MFIP_CHECKBOX = CHECKED
	END IF
	IF cash2_status_check = "INAC" or cash2_status_check = "SUSP" or cash2_status_check = "DENY" or cash2_status_check = "" THEN MF_STATUS = FALSE
END IF

IF cash1_prog_check = "DW" THEN
	IF cash1_status_check = "ACTV" or cash1_status_check = "PEND" or cash2_status_check = "ACTV" or cash2_status_check = "PEND" THEN
		DWP_STATUS = TRUE
		DWP_CHECKBOX  = CHECKED
	END IF
	If cash1_status_check = "INAC" or cash1_status_check = "SUSP" or cash1_status_check = "DENY" or cash1_status_check = "" THEN DWP_STATUS = FALSE
	If cash2_status_check = "INAC" or cash2_status_check = "SUSP" or cash2_status_check = "DENY" or cash2_status_check = "" THEN DWP_STATUS = FALSE
END IF

If cash1_prog_check = "" THEN
	If cash1_status_check = "PEND" or cash2_status_check = "PEND" THEN
		CASH_STATUS = TRUE
		CASH_CHECKBOX = CHECKED
	END IF
	If cash1_status_check = "INAC" or cash1_status_check = "SUSP" or cash1_status_check = "DENY" or cash1_status_check = "" THEN CASH_STATUS = FALSE
END IF

If cash2_prog_check = "" THEN
	If cash2_status_check = "INAC" or cash2_status_check = "SUSP" or cash2_status_check = "DENY" or cash2_status_check = "" THEN CASH_STATUS = FALSE
END IF

IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN GRH_STATUS = TRUE
'can you say and or
IF active_status = FALSE THEN
 	IF MF_STATUS = FALSE and FS_STATUS = FALSE and HC_STATUS = FALSE and DWP_STATUS = FALSE and CASH_STATUS = FALSE THEN
		case_note_only = TRUE
		msgbox "It appears no HC, FS, or Cash are open on this case."
	END IF
END IF

'----------------------------------------------------------------------------------------------------DIALOGS
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 376, 285, "Good Cause"
  EditBox 70, 5, 30, 15, child_ref_number
  EditBox 70, 25, 40, 15, actual_date
  EditBox 70, 45, 40, 15, claim_date
  EditBox 70, 65, 40, 15, review_date
  DropListBox 165, 5, 70, 15, "Select One:"+chr(9)+"Denied"+chr(9)+"Granted"+chr(9)+"Not Claimed"+chr(9)+"Pending", gc_status
  DropListBox 165, 25, 115, 15, "Select One:"+chr(9)+"Application Review-Complete"+chr(9)+"Application Review-Incomplete"+chr(9)+"Change/exemption ending"+chr(9)+"Determination"+chr(9)+"Recertification", good_cause_droplist
  DropListBox 165, 45, 115, 15, "Select One:"+chr(9)+"Potential phys harm/child"+chr(9)+"Potential emotnl harm/child"+chr(9)+"Potential phys harm/caregiver"+chr(9)+"Potential emotnl harm/caregiver"+chr(9)+"Cncptn incest/forced rape"+chr(9)+"Legal adoption before court"+chr(9)+"Parent gets preadoptn svc"+chr(9)+"No longer claiming", reason_droplist
  EditBox 165, 65, 205, 15, denial_reason
  CheckBox 295, 15, 25, 10, "CCA", CCA_CHECKBOX
  CheckBox 335, 15, 25, 10, "HC", HC_CHECKBOX
  CheckBox 295, 30, 30, 10, "DWP", DWP_CHECKBOX
  CheckBox 335, 30, 20, 10, "MF", MFIP_CHECKBOX
  CheckBox 295, 45, 20, 10, "FS", FS_CHECKBOX
  CheckBox 335, 45, 30, 10, "METS", METS_CHECKBOX
  CheckBox 10, 95, 145, 10, "ABPS name not written on the correct line", ABPS_CHECKBOX
  CheckBox 160, 95, 120, 10, "All of the questions not answered", QUESTIONS_CHECKBOX
  CheckBox 285, 95, 80, 10, "Other (please specify)", OTHER_CHECKBOX
  CheckBox 10, 105, 140, 10, "Reason for requesting GC not selected", REASON_CHECKBOX
  CheckBox 160, 105, 90, 10, "No signature and/or date ", NOSIG_CHECKBOX
  EditBox 65, 125, 85, 15, mets_info
  EditBox 220, 125, 150, 15, verifs_req
  EditBox 65, 145, 305, 15, other_notes
  CheckBox 295, 175, 55, 10, "Investigation", investigation_check
  CheckBox 295, 185, 60, 10, "Sup Evidence", sup_evidence_check
  CheckBox 295, 195, 70, 10, "Med Sup Svc Only", med_sup_check
  CheckBox 10, 175, 185, 10, "Request for Proof to Support Good Cause Claim", SUP_CHECKBOX
  CheckBox 10, 185, 160, 10, "Good Cause Client Statement (DHS-2338)", DHS_2338_CHECKBOX
  CheckBox 10, 195, 220, 10, "Imp Information about Your Request Exemption (DHS-3627) ", DHS_3627_CHECKBOX
  CheckBox 10, 205, 205, 10, "Notice of Denial of Good Cause Exemption (DHS-3628) ", DHS_3628_CHECKBOX
  CheckBox 10, 215, 165, 10, "Notice of Good Cause Approval (DHS-3629) ", DHS_3629_CHECKBOX
  CheckBox 10, 225, 195, 10, "Request to End Good Cause Exemption  (DHS-3631 )", DHS_3631_CHECKBOX
  CheckBox 10, 235, 180, 10, "Request for Additional Information (DHS 3632)", DHS_3632_CHECKBOX
  CheckBox 10, 245, 165, 10, "Good Cause Yearly Determination Packet", Recert_CHECKBOX
  CheckBox 10, 255, 195, 10, "Good Cause Redetermination Approval ( DHS 3633)", DHS_3633_CHECKBOX
  CheckBox 10, 265, 245, 10, "Good Cause Client Statement (DHS-2338) is in ECF and completed in full", DHS_2338_complete_CHECKBOX
  ButtonGroup ButtonPressed
    OkButton 265, 265, 50, 15
    CancelButton 320, 265, 50, 15
  Text 25, 30, 40, 10, "Actual date:"
  Text 140, 30, 25, 10, "Action:"
  Text 135, 50, 30, 10, "Reason:"
  Text 140, 10, 25, 10, "Status:"
  Text 5, 130, 60, 10, "Mets information:"
  Text 160, 130, 60, 10, "Verifs requested:"
  Text 115, 70, 50, 10, "Denial reason:"
  Text 20, 150, 45, 10, "Other notes:"
  Text 25, 70, 45, 10, "Next review:"
  GroupBox 285, 165, 85, 45, "Follow up"
  GroupBox 5, 85, 365, 35, "Incomplete Form:"
  Text 25, 50, 40, 10, "Claim date:"
  Text 5, 10, 65, 10, "Child's MEMB #(s):"
  GroupBox 5, 165, 255, 115, "Verifications"
  GroupBox 290, 5, 80, 55, "Programs"
EndDialog

Do
	Do
		err_msg = ""
		dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF child_ref_number = "" THEN err_msg = err_msg & vbnewline & "* Please enter the child(s) member number."
		If good_cause_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the Good Cause Application Status."
		IF gc_status = "Select One:" THEN err_msg = err_msg & vbnewline & "* Select the Good Cause Case Status."
		If reason_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the Good Cause reason."
		If gc_status = "Granted" THEN
			If isdate(review_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid good cause review date."
		ELSEIF gc_status = "Denied" THEN
			If denial_reason = "" then err_msg = err_msg & vbnewline & "* Please enter a denial reason."
		END IF
		If isdate(actual_date) = FALSE THEN
			err_msg = err_msg & vbnewline & "* Please enter a date."
		Else
		 	IF cdate(actual_date) > cdate(date) = TRUE THEN err_msg = err_msg & vbnewline & "* Please enter an actual date that is not in the future and is in the footer month that you are working in."
		END IF
		IF gc_status <> "Not Claimed" THEN
			If isdate(claim_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid good cause claim date."
		END IF
		If good_cause_droplist = "Application Review-Incomplete" then
			If other_notes = "" and OTHER_CHECKBOX = CHECKED then err_msg = err_msg & vbnewline & "* Please a reason that application is incomplete."
		END IF
		IF METS_CHECKBOX = CHECKED and mets_info = "" THEN err_msg = err_msg & vbnewline & "* Please enter a METS case number or unknown."
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'----------------------------------------------------------------------------------------------------ABPS panel
MAXIS_background_check
Call MAXIS_footer_month_confirmation
'build in something to confimr claim date vs footer month'
Call navigate_to_MAXIS_screen("STAT", "ABPS")
'Initial dialog giving the user the option to select the type of good cause action
'Checks to make sure there are ABPS panels for this member. If none exist the script will close
EMReadScreen total_amt_of_panels, 1, 2, 78
If total_amt_of_panels = "0" then script_end_procedure_with_error_report("An ABPS panel does not exist. Please create the panel before running the script again.")
EMReadScreen panel_check, 4, 2, 50
'If panel_check = "ABPS" and current_panel_number = total_amt_of_panels then
IF panel_check = "ABPS" THEN
	'MsgBox panel_check
    Do
		EMReadScreen initial_parental_status_number, 1, 15, 53
		IF initial_parental_status_number = "1" THEN
			EMReadScreen ABPS_last_name, 24, 10, 30	'reading last name
			ABPS_last_name = replace(ABPS_last_name, "_", "")
			'MsgBox ABPS_last_name
			EMReadScreen ABPS_first_name, 12, 10, 63	'reading first name
			ABPS_first_name = replace(ABPS_first_name, "_", "")
		END IF
		EMReadScreen ABPS_parent_ID, 10, 13, 40	'making sure ABPS is not unknown.
		ABPS_parent_ID = trim(ABPS_parent_ID)
       	EMReadScreen current_panel_number, 1, 2, 73
		'MsgBox current_panel_number
       	ABPS_check = MsgBox("Is this the correct ABPS panel to update?  " & ABPS_first_name & " " & ABPS_last_name & " ID# " &  ABPS_parent_ID, vbYesNo + vbQuestion, "Initial Confirmation")
       	If ABPS_check = vbYes then
			ABPS_found = TRUE
			exit do
		END IF
       	If ABPS_check = vbNo then
			ABPS_found = FALSE
			TRANSMIT
		END IF
		If (ABPS_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure_with_error_report("Unable to find another ABPS. Please review the case, and run the script again if applicable.")
    Loop until ABPS_found = TRUE
END IF

'-------------------------------------------------------------------------Updating the ABPS panel
EMReadScreen MAXIS_footer_month, 2, 20, 55
'MsgBox MAXIS_footer_month
DO
    PF9'edit mode
	DO
		EmReadscreen edit_mode_check, 1, 20, 08
		IF edit_mode_check = "D" THEN
        	PF9
			'msgbox "are we in the edit mode"
		ElseIF edit_mode_check = "E" THEN
			EXIT DO
		END IF
	Loop Until edit_mode_check = "D"

	EMReadScreen parental_status_number, 1, 15, 53	'making sure ABPS is not unknown.
	'MsgBox parental_status_number
	EMReadScreen custodial_status, 1, 15, 67
	'MsgBox custodial_status
	IF parental_status_number = "1" THEN
		EMReadScreen first_name, 12, 10, 63
		EMReadScreen last_name, 24, 10, 30
	    first_name = trim(first_name)
	    last_name = trim(last_name)
	    first_name = replace(first_name, "_", "")
	    last_name = replace(last_name, "_", "")
	    client_name = first_name & " " & last_name
	END IF
	EMReadScreen ABPS_gender, 1, 11, 80	'reading the ssn
	EMReadScreen ABPS_SSN, 11, 11, 30	'reading the ssn
	ABPS_SSN = replace(ABPS_SSN, " ", "-")
	EMReadScreen ABPS_DOB, 10, 11, 60	'reading the DOB
	ABPS_DOB = replace(ABPS_DOB, " ", "-")
	EMReadScreen HC_ins_order, 1, 12, 44	'making sure ABPS is not unknown.
	EMReadScreen HC_ins_compliance, 1, 12, 80

	IF parental_status_number = "1" THEN parental_status = "Absent parent Known/Alleged"
	IF parental_status_number = "2" THEN parental_status = "Absent parent Unknown"
    IF parental_status_number = "3" THEN parental_status = "Absent parent Deceased"
    IF parental_status_number = "4" THEN parental_status = "Parental rights severed"
    IF parental_status_number = "5" THEN parental_status = "N/A, minor is non-Unit mbr"
    IF parental_status_number = "6" THEN parental_status = "Minor crgvr no order sup"
    IF parental_status_number = "7" THEN parental_status = "Appl/HC child no order sup"

	'MsgBox "Case Note Only: " & case_note_only
	IF edit_mode_check = "E" THEN  'msgbox "what are we doing"
		case_note_only = FALSE
		EMWriteScreen "Y", 4, 73			'Support Coop Y/N field
		'msgbox "should be writing"
	    IF gc_status = "Pending" THEN
	    	'Msgbox gc_status
	    	EMWriteScreen "P", 5, 47			'Good Cause status field
	    	EMWriteScreen "  ", 6, 73'next review date'
	    	EMWriteScreen "  ", 6, 76'next review date'
	    	EMWriteScreen "  ", 6, 79'next review date'
	    ELSEIF gc_status = "Granted" THEN
	    	'Msgbox gc_status
	    	EMWriteScreen "G", 5, 47
	    	Call create_MAXIS_friendly_date(datevalue(review_date), 0, 6, 73)
	    ELSEIF gc_status = "Denied" THEN
	    	'Msgbox gc_status
	    	EMWriteScreen "D", 5, 47
	    	EMWriteScreen "  ", 6, 73'next review date'
	    	EMWriteScreen "  ", 6, 76'next review date'
	    	EMWriteScreen "  ", 6, 79'next review date'
	    ELSEIF gc_status = "Not Claimed" THEN
	    	'Msgbox gc_status
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
	    If reason_droplist = "Potential phys harm/child"		then claim_reason = "1"
	    If reason_droplist = "Potential emotnl harm/child"	 	then claim_reason = "2"
	    If reason_droplist = "Potential phys harm/caregiver" 	then claim_reason = "3"
	    If reason_droplist = "Potential emotnl harm/caregiver" 	then claim_reason = "4"
	    If reason_droplist = "Cncptn incest/forced rape" 		then claim_reason = "5"
	    If reason_droplist = "Legal adoption before court" 		then claim_reason = "6"
	    If reason_droplist = "Parent gets preadoptn svc" 		then claim_reason = "7"

	    IF gc_status <> "Not Claimed" THEN
	    	Call create_MAXIS_friendly_date(datevalue(claim_date), 0, 5, 73)
	    	IF sup_evidence_CHECKBOX = CHECKED THEN EMWriteScreen "Y", 7, 47 ELSE EMWriteScreen "N", 7, 47
	    	IF investigation_CHECKBOX = CHECKED THEN EMWriteScreen "Y", 7, 73 ELSE EMWriteScreen "N", 7, 73
	    	IF med_sup_CHECKBOX = CHECKED THEN EMWriteScreen "Y", 8, 48 ELSE EMWriteScreen "N", 8, 48
	    	EMWriteScreen claim_reason, 6, 47
	    	Call create_MAXIS_friendly_date_with_YYYY(datevalue(actual_date), 0, 18, 38) 'creates and writes the date entered in dialog'
	    END IF
		'msgbox "did we make it to actual date"
	    Call create_MAXIS_friendly_date_with_YYYY(datevalue(actual_date), 0, 18, 38) 'creates and writes the date entered in dialog'
		TRANSMIT  'to add information
		' CASE STATUS IS 'INACTIVE' FOR THIS PERIOD. NO BACKGROUND TRANSACTIONS CREATED  '
		EmReadscreen ABPS_error_check, 50, 24, 2
		IF trim(ABPS_error_check) <> "" THEN
			TRANSMIT 'this will get passed inhibiting errors''
		ELSE
			IF ABPS_error_check <> "" THEN MsgBox "*** Error Check ***" & vbNewLine & ABPS_error_check & vbNewLine
			'If ABPS_error_check = "GOOD CAUSE CLAIM DATE CANNOT BE IN THE FUTURE" THEN script_end_procedure_with_error_report(ABPS_error_check & vbNewLine & "Please run the script again.")
		END IF

		EMReadScreen panel_check, 4, 2, 50
		IF panel_check = "ABPS" THEN
			CALL write_value_and_transmit("BGTX", 20, 71)
    		'msgbox "now i want to be at wrap"
    		EMReadScreen WRAP_panel_check, 4, 2, 46
    	    IF WRAP_panel_check = "WRAP" THEN
				EMReadScreen MAXIS_footer_month, 2, 20, 55
            	'msgbox "Footer month: " & MAXIS_footer_month & vbcr & "CM Plus one" & CM_plus_1_mo
            	IF MAXIS_footer_month <> CM_plus_1_mo THEN
                    Call write_value_and_transmit("Y", 16, 54)
					EmReadscreen bgtx_error, 4, 24, 15
					IF bgtx_error = "BGTX" THEN EXIT DO
                	EMReadScreen check_PNLP, 4, 2, 53
                    'msgbox "check PNLP: " & check_PNLP
                    IF check_PNLP = "PNLP" THEN
				    	CALL write_value_and_transmit("ABPS", 20, 71)
				    	Do
				    		'DO
                                IF parental_status = "1" THEN
				    	            EMReadScreen ABPS_last_name_check, 24, 10, 30	'reading last name
				    	            ABPS_last_name_check = replace(ABPS_last_name_check, "_", "")
				    	            MsgBox ABPS_last_name_check & "second run"
				    	            EMReadScreen ABPS_first_name_check, 12, 10, 63	'reading first name
				    	            ABPS_first_name_check = replace(ABPS_first_name_check, "_", "")
				    			END IF
				    	        EMReadScreen ABPS_parent_ID, 10, 13, 40	'making sure ABPS is not unknown.
				    	        ABPS_parent_ID = trim(ABPS_parent_ID)
			           	        EMReadScreen current_panel_number, 1, 2, 73
				    	        'MsgBox current_panel_number
				    			'IF ABPS_last_name = ABPS_last_name_check and ABPS_first_name = ABPS_first_name_check THEN
				    			'	ABPS_found = TRUE
				    			'	MsgBox ABPS_found
				    			'	'exit do
				    			'ELSE
				    			'	ABPS_found = FALSE
				    			'	TRANSMIT
				    			'	MsgBox ABPS_found
				    			'END IF
				    			'IF ABPS_found = TRUE THEN
				    				ABPS_check = MsgBox("Is this the correct ABPS panel to update? Sometimes a new ID is created for the same ABPS.  " & ABPS_first_name_check & " " & ABPS_last_name_check & " ID# " &  ABPS_parent_ID, vbYesNo + vbQuestion, "BGTX Confirmation")
				    				If ABPS_check = vbYes then
				    					ABPS_found_question = TRUE
				    					exit do
				    				END IF
				    				If ABPS_check = vbNo then
				    					ABPS_found_question = FALSE
				    					TRANSMIT
				    				END IF
				    	    		If (ABPS_check = vbNo AND current_panel_number = total_amt_of_panels) THEN script_end_procedure_with_error_report("Unable to find another ABPS. Please review the case, and run the script again if applicable.")
				    			'END IF
			        		'Loop until ABPS_found = TRUE
				    	Loop until ABPS_found_question = TRUE
                    END IF
				END IF
			END IF
	    ELSE
			MsgBox "*** NOTICE!!! ***" & vbNewLine & "Unable to complete action. " & vbNewLine
		END IF
	END IF
LOOP Until MAXIS_footer_month = CM_plus_1_mo

If good_cause_droplist = "Change/exemption ending" then
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 216, 100, "Good cause change/exemption "
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
	Do
		Do
			err_msg = ""
			dialog Dialog1
			cancel_confirmation
			If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

'-----------------------------------------------------------------------------------------For case note'
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

'Call MAXIS_footer_month_confirmation    'Footer month & year could get wonky and not go into case note. This prevents that from happening.

'original_MAXIS_footer_month = MAXIS_footer_month
MAXIS_footer_month = original_MAXIS_footer_month
'MsgBox MAXIS_footer_month
'msgbox original_MAXIS_footer_month

Call navigate_to_MAXIS_screen("CASE", "NOTE")
PF9
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
Call write_variable_in_case_note("* Child(ren) member number(s): " & child_ref_number)
Call write_bullet_and_variable_in_case_note("ABPS name", client_name)
Call write_bullet_and_variable_in_case_note("Parental status", parental_status)
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
IF Recert_CHECKBOX   = CHECKED THEN Call write_variable_in_case_note("* Sent Good Cause Yearly Determination Packet")
IF DHS_3633_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Sent Good Cause Redetermination Approval (DHS 3633)")
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)
'PF3
IF mets_info = "" THEN
	IF FS_CHECKBOX = CHECKED and CASH_CHECKBOX = UNCHECKED and CCA_CHECKBOX = UNCHECKED and DWP_CHECKBOX = UNCHECKED and MFIP_CHECKBOX = UNCHECKED and HC_CHECKBOX = UNCHECKED and METS_CHECKBOX = UNCHECKED THEN memo_started = TRUE
END IF
IF memo_started = TRUE THEN
	Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    
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
PF3
Script_end_procedure_with_error_report("Success! MAXIS has been updated, and the Good Cause results case noted.")
