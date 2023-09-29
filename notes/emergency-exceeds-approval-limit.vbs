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
    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, EMER_type, case_status, list_active_programs, list_pending_programs)
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
	Do
		err_msg = ""

		dialog Diaog1
		cancel_without_confirmation

       Call validate_MAXIS_case_number(err_msg, "*")
	   If IsNumeric(chck_approval_limit) = False Then err_msg = err_msg & vbCr & "* Enter the max amount to are authorized to approve."
	   If trim(worker_signature) = "" Then err_msg = err_msg & vbCr & "* Sign your CASE/NOTE by entering your name in the 'Worker Signature' field."

	   If err_msg <> "" Then MsgBox "*  *  *  NOTICE  *  *  *" & vbCr & "Please resolve the following to continue:" & vbCr & err_msg
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False
Call Check_for_MAXIS(False)


'dialog to select emergeny types and record check specifics.


'possibly email?


'CASE/NOTE of the information about the case approval



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




