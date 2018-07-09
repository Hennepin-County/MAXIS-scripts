'GATHERING STATS===========================================================================================
name_of_script = "NOTES - OVERPAYMENT CLAIM ENTERED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
CALL changelog_update("05/01/2018", "Updated script to ensure Reason for OP is entered as it is a mandatory field.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/25/2018", "Updated script to add Fraud and Earned Income handling.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Updated script to add HC handling and the income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/01/2018", "Updated script to write amount in case note in the correct area.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/04/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


function DEU_password_check(end_script)
'--- This function checks to ensure the user is in a MAXIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MAXIS screen.
'===== Keywords: MAXIS, production, script_end_procedure
	Do
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
			If end_script = True then
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
				warning_box = MsgBox("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

EMConnect ""

CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
OP_Date = date & ""

BeginDialog match_claim_dialog, 0, 0, 361, 245, "Overpayment Claim Entered"
	EditBox 55, 5, 35, 15, MAXIS_case_number
	EditBox 150, 5, 45, 15, discovery_date
	EditBox 240, 5, 20, 15, memb_number
  EditBox 330, 5, 20, 15, OT_resp_memb
	DropListBox 55, 25, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
  EditBox 130, 65, 30, 15, OP_from
  EditBox 180, 65, 30, 15, OP_to
  EditBox 245, 65, 35, 15, Claim_number
  EditBox 305, 65, 45, 15, Claim_amount
  DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
  EditBox 130, 85, 30, 15, OP_from_II
  EditBox 180, 85, 30, 15, OP_to_II
  EditBox 245, 85, 35, 15, Claim_number_II
  EditBox 305, 85, 45, 15, Claim_amount_II
  DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
  EditBox 130, 105, 30, 15, OP_from_III
  EditBox 180, 105, 30, 15, OP_to_III
  EditBox 245, 105, 35, 15, Claim_number_III
  EditBox 305, 105, 45, 15, Claim_amount_III
  EditBox 70, 140, 160, 15, EVF_used
	EditBox 305, 140, 50, 15, income_rcvd_date
	EditBox 70, 160, 285, 15, Reason_OP
	DropListBox 105, 180, 35, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", EI_allowed_dropdown
	DropListBox 185, 180, 35, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  DropListBox 255, 180, 100, 15, "Select:"+chr(9)+"HH No Info"+chr(9)+"HH Incorrect Or Incompl"+chr(9)+"HH Not Timely Inform"+chr(9)+"Estimate Wrong Inadv"+chr(9)+"No HRF"+chr(9)+"Benefits Rcd Pend App"+chr(9)+"Replacement Bene Used" +chr(9)+"Prog Pol Prevents Chg" +chr(9)+"EBT- Representment" +chr(9)+"IEVS - BEER"+chr(9)+"IEVS - BENDEX" +chr(9)+"IEVS - UNVI"+chr(9)+"IEVS - SDX"+chr(9)+"IEVS - WAGE"+chr(9)+"IEVS - UBEN"+chr(9)+"PARIS Inter-state Match"+chr(9)+"Agency: Delay   Action"+chr(9)+"Agency: Issue Comp Error"+chr(9)+"Agency: Dup Issuance"+chr(9)+"Agency:$50 Pass-Thru"+chr(9)+"GRH Vndr: No Info"+chr(9)+"GRH Vndr: Incomplete"+chr(9)+"GRH Vndr: Not Timely"+chr(9)+"GRH Vndr: Client Left Faci"+chr(9)+"99 Other", collectible_reason_dropdown
	EditBox 95, 200, 60, 15, HC_resp_memb
  EditBox 290, 200, 65, 15, Fed_HC_AMT
  ButtonGroup ButtonPressed
      OkButton 260, 225, 45, 15
      CancelButton 310, 225, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 95, 10, 55, 10, "Discovery Date: "
  Text 5, 30, 50, 10, "Fraud referral:"
  Text 205, 10, 30, 10, "MEMB #:"
  Text 270, 10, 60, 10, "Other resp. memb:"
  GroupBox 10, 45, 345, 90, "Overpayment Information"
  Text 15, 70, 30, 10, "Program:"
  Text 105, 70, 20, 10, "From:"
  Text 165, 70, 10, 10, "To:"
  Text 215, 70, 25, 10, "Claim #"
  Text 285, 70, 20, 10, "AMT:"
  Text 15, 90, 30, 10, "Program:"
  Text 105, 90, 20, 10, "From:"
  Text 165, 90, 10, 10, "To:"
  Text 215, 90, 25, 10, "Claim #"
  Text 285, 90, 20, 10, "AMT:"
  Text 15, 110, 30, 10, "Program:"
  Text 105, 110, 20, 10, "From:"
  Text 165, 110, 10, 10, "To:"
  Text 215, 110, 25, 10, "Claim #"
  Text 285, 110, 20, 10, "AMT:"
	Text 5, 185, 100, 10, "E. Income disregard allowed?"
  Text 145, 185, 40, 10, "Collectible?"
  Text 225, 185, 30, 10, "Reason:"
  Text 5, 145, 60, 10, "Income verif used:"
  Text 05, 205, 90, 10, "HC responsible members:"
  Text 200, 205, 85, 10, "Total Federal HC amount:"
  Text 15, 165, 50, 10, "Reason for OP:"
  Text 240, 145, 60, 10, "Date income rcvd: "
	Text 180, 55, 30, 10, "(MM/YY)"
  Text 130, 55, 30, 10, "(MM/YY)"
EndDialog

Do
	err_msg = ""
	dialog match_claim_dialog
	IF buttonpressed = 0 then stopscript
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF trim(Reason_OP) = "" or len(Reason_OP) < 8 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 8)."
	IF OP_program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
	IF OP_program_II <> "Select:" THEN
		IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	END IF
	IF OP_program_III <> "Select:" THEN
		IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
		IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
		IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	END IF
	IF IEVS_type = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
	IF EI_allowed_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if Earned Income disregard was allowed."
  IF collectible_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if claim is collectible."
	IF collectible_dropdown = "YES" THEN
	IF collectible_reason_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise why claim is collectible."
	END IF
	IF income_rcvd_date = "" THEN err_msg = err_msg & vbnewline & "* Please advise of date income was received."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL DEU_password_check(False)

'----------------------------------------------------------------------------------------------------STAT
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
EMReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
first_name = trim(first_name)
'-----------------------------------------------------------------------------------------CASENOTE
'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("OVERPAYMENT CLAIM ENTERED" & " (" & first_name & ") " & OP_from & " through " & OP_to)
CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned Income Disregard Allowed", EI_allowed_dropdown)
IF OP_program = "HC" THEN
	Call write_bullet_and_variable_in_CASE_NOTE("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CASE_NOTE("Total federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
CALL write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_case_note("Date income verification was received", income_rcvd_date)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_case_note("MANDATORY-Reason for overpayment", Reason_OP)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3
IF OP_program = "HC" THEN CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us", "Claim entered for case #" &  MAXIS_case_number, "Member #: " & OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount & "See case notes for further details.", "", False)

script_end_procedure("Overpayment case note entered. Please remember to copy and paste your notes to CCOL/CLIC")
