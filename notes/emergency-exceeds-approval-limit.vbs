'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EMERGENCY EXCEEDS APPROVAL LIMIT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
call changelog_update("10/1/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================
'END FUNCTIONS BLOCK =======================================================================================================


'SCRIPT ====================================================================================================================
EMConnect "" 								'Connect to MAXIS
Call Check_for_MAXIS(False)
Call MAXIS_case_number_finder(MAXIS_case_number)				'Capture CASE/NUMBER

If MAXIS_case_number <> "" Then 								'Try to identify if the case is EA or EGA
	EMReadScreen on_mony_chck, 4, 2, 46
	If on_mony_chck = "CHCK" Then
		EMReadScreen EMER_type, 2, 5, 17
		If EMER_type = "EG" Then EMER_type = "EGA"
		'READ CHCK INFORMAITON
	Else
    	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, EMER_type, case_status, list_active_programs, list_pending_programs)
	End If
End If

chck_approval_limit = "1200" 									'default the approval limit to 1200 for EA
If EMER_type = "EGA" then chck_approval_limit = "4000"			'default the approval limit to 4000 for EGA

'dialog to confirm Case Number, EMER program, and Worker Signature
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 216, 185, "Dialog"
	EditBox 95, 15, 50, 15, MAXIS_case_number
	DropListBox 95, 35, 60, 45, "Select One..."+chr(9)+"EA"+chr(9)+"EGA", EMER_type
	EditBox 95, 55, 50, 15, chck_approval_limit
	EditBox 10, 165, 140, 15, worker_signature
	ButtonGroup ButtonPressed
		OkButton 160, 145, 50, 15
		CancelButton 160, 165, 50, 15
	GroupBox 10, 5, 195, 70, "Case Information"
	Text 45, 20, 50, 10, "Case Number:"
	Text 20, 40, 70, 10, "Emergency Program:"
	Text 25, 60, 70, 10, "Your approval limit:"
	Text 10, 80, 100, 10, "What is this script used for?"
	Text 15, 90, 190, 25, "Purpose: Document details of emergency need/issuance for cases that require more funds than you are authorized to approve."
	Text 10, 120, 75, 10, "What the script does:"
	Text 15, 130, 135, 20, "Creates a CASE/NOTE with all the details for the check issuance."
	Text 10, 155, 65, 10, "Worker Signature:"
EndDialog

Do
	err_msg = ""

	dialog Diaog1
	cancel_without_confirmation

	Call validate_MAXIS_case_number(err_msg, "*")
	If IsNumeric(chck_approval_limit) = False Then err_msg = err_msg & vbCr & "* Enter the max amount to are authorized to approve."
	If trim(worker_signature) = "" Then err_msg = err_msg & vbCr & "* Sign your CASE/NOTE by entering your name in the 'Worker Signature' field."

	If err_msg <> "" Then MsgBox "*  *  *  NOTICE  *  *  *" & vbCr & "Please resolve the following to continue:" & vbCr & err_msg
Loop until err_msg = ""

reason_code_list = "Select or Type"
reason_code_list = reason_code_list+chr(9)+"03 Home Repair"
reason_code_list = reason_code_list+chr(9)+"04 HH Furnishings"
reason_code_list = reason_code_list+chr(9)+"10 Transportation"
reason_code_list = reason_code_list+chr(9)+"12 Other"
reason_code_list = reason_code_list+chr(9)+"26 Shelter Not FV"
reason_code_list = reason_code_list+chr(9)+"28 Utility Shut-off"
reason_code_list = reason_code_list+chr(9)+"29 Foreclosure"
reason_code_list = reason_code_list+chr(9)+"30 Moving Exp"
reason_code_list = reason_code_list+chr(9)+"35 Temporary Housing"
reason_code_list = reason_code_list+chr(9)+"40 Damage Deposit"
reason_code_list = reason_code_list+chr(9)+"44 Permanent Housing"

add_check_2 = False
approval_date = date & ""

'dialog to select emergeny types and record check specifics.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 466, 205, "Emergency Issuance Details"
	ButtonGroup ButtonPressed
	EditBox 410, 15, 40, 15, approval_date
	ComboBox 15, 95, 130, 45, reason_code_list+chr(9)+check_1_reason, check_1_reason
	EditBox 155, 95, 40, 15, check_1_amount
	EditBox 210, 95, 30, 15, check_1_start_date
	EditBox 245, 95, 30, 15, check_1_end_date
	EditBox 285, 95, 25, 15, check_1_elig_hh_membs
	EditBox 320, 95, 35, 15, check_1_vendor
	EditBox 365, 95, 75, 15, check_1_client_ref
	If add_check_2 = True Then
		ComboBox 15, 115, 130, 45, reason_code_list+chr(9)+check_2_reason, check_2_reason
		EditBox 155, 115, 40, 15, check_2_amount
		EditBox 210, 115, 30, 15, check_2_start_date
		EditBox 245, 115, 30, 15, check_2_end_date
		EditBox 285, 115, 25, 15, check_2_elig_hh_membs
		EditBox 320, 115, 35, 15, check_2_vendor
		EditBox 365, 115, 75, 15, check_2_client_ref
	Else
		PushButton 340, 120, 115, 15, "This case requires two checks", add_another_check
	End If
	EditBox 5, 160, 455, 15, emer_approval_notes
	OkButton 355, 185, 50, 15
	CancelButton 410, 185, 50, 15
  	GroupBox 5, 5, 455, 30, "Case Information"
  	Text 15, 20, 110, 10, "Case Number:"
  	Text 130, 20, 90, 10, "Emergency Program:"
  	Text 235, 20, 100, 10, "Your approval limit:"
  	Text 355, 20, 50, 10, "Approval Date:"
	Text 10, 40, 300, 10, "Enter all information needed to create the check to make issuance easier for the approver."
	Text 10, 50, 215, 10, "This script currently allows for 2 unique checks to be entered. "
	GroupBox 5, 70, 455, 70, "Checks Needed to Resolve Emergency:"
	Text 15, 85, 65, 10, "Check Reason"
	Text 165, 85, 50, 10, "Check Amount"
	Text 230, 85, 50, 10, "Period"
	Text 320, 85, 65, 10, "Elig HH Membs"
	Text 320, 85, 30, 10, "Vendor"
	Text 365, 85, 50, 10, "Client Ref Number:"
	Text 260, 97, 5, 10, "--"
	Text 5, 150, 120, 10, "Emergency Determination Notes:"
EndDialog

Do
	Do
		err_msg = ""

		dialog Diaog1
		cancel_confirmation


		If ButtonPressed = add_another_check Then
			err_msg = "LOOP" & err_msg
			add_check_2 = True
		End If

		If err_msg <> "" and left(err_msg, 4) <> "LOOP" Then MsgBox "*  *  *  NOTICE  *  *  *" & vbCr & "Please resolve the following to continue:" & vbCr & err_msg
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False
Call Check_for_MAXIS(False)

'possibly email?


'CASE/NOTE of the information about the case approval
Call start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE(EMER_typt & "Determination Done - Check Issuance Still Needed")
Call write_variable_in_CASE_NOTE("Approval of " & EMER_type & " has been compelted, but issuance is restricted and a separate approver is needed to issue the payment.")
Call write_bullet_and_variable_in_CASE_NOTE ("Approval Date", approval_date)
Call write_variable_in_CASE_NOTE("Check Details ===============================================")
If add_check_2 = True Then Call write_variable_in_CASE_NOTE("CHECK ONE")
Call write_variable_in_CASE_NOTE("   Amount: " & check_1_amount)
Call write_variable_in_CASE_NOTE("   Period: " & check_1_start_date & " - " & check_1_end_date)
Call write_variable_in_CASE_NOTE("   ELIG HH Members: " & check_1_elig_hh_membs)
Call write_variable_in_CASE_NOTE("   Vendor: " & check_1_vendor)
Call write_variable_in_CASE_NOTE("   Client Ref Number: " & check_1_client_ref)
If add_check_2 = True Then
	Call write_variable_in_CASE_NOTE("CHECK TWO")
	Call write_variable_in_CASE_NOTE("   Amount: " & check_2_amount)
	Call write_variable_in_CASE_NOTE("   Period: " & check_2_start_date & " - " & check_2_end_date)
	Call write_variable_in_CASE_NOTE("   ELIG HH Members: " & check_2_elig_hh_membs)
	Call write_variable_in_CASE_NOTE("   Vendor: " & check_2_vendor)
	Call write_variable_in_CASE_NOTE("   Client Ref Number: " & check_2_client_ref)
End If
Call write_variable_in_CASE_NOTE("=============================================================")
Call write_bullet_and_variable_in_CASE_NOTE("Note", emer_approval_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)


'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------




