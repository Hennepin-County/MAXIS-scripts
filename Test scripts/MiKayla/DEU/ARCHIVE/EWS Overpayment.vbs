'GATHERING STATS===========================================================================================
name_of_script = "OVERPAYMENT CLAIM ENTERED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

BeginDialog EWS_OP_dialog, 0, 0, 396, 205, "Overpayment Claim Entered"
  EditBox 55, 5, 35, 15, MAXIS_case_number
  EditBox 130, 5, 20, 15, memb_number
  EditBox 225, 5, 20, 15, OT_resp_memb
  EditBox 310, 5, 70, 15, Discovery_date
  DropListBox 45, 45, 50, 15, "Select:"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"DW", First_Program
  EditBox 125, 45, 20, 15, First_from_IEVS_month
  EditBox 155, 45, 20, 15, First_from_IEVS_year
  EditBox 195, 45, 20, 15, First_to_IEVS_month
  EditBox 220, 45, 20, 15, First_to_IEVS_year
  EditBox 275, 45, 40, 15, First_OP
  EditBox 340, 45, 40, 15, First_AMT
  DropListBox 45, 65, 50, 15, "Select:"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"DW", Second_Program
  EditBox 125, 65, 20, 15, Second_from_IEVS_month
  EditBox 155, 65, 20, 15, Second_from_IEVS_year
  EditBox 195, 65, 20, 15, Second_to_IEVS_month
  EditBox 220, 65, 20, 15, Second_to_IEVS_year
  EditBox 275, 65, 40, 15, Second_OP
  EditBox 340, 65, 40, 15, Second_AMT
  DropListBox 45, 85, 50, 15, "Select:"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"DW", Third_Program
  EditBox 125, 85, 20, 15, Third_from_IEVS_month
  EditBox 155, 85, 20, 15, Third_from_IEVS_year
  EditBox 195, 85, 20, 15, Third_to_IEVS_month
  EditBox 220, 85, 20, 15, Third_from_IEVS_year
  EditBox 275, 85, 40, 15, Third_OP
  EditBox 340, 85, 40, 15, Third_AMT
  DropListBox 50, 120, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  EditBox 165, 120, 120, 15, collectible_reason
  DropListBox 340, 120, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  EditBox 60, 140, 80, 15, source_income
  EditBox 235, 140, 145, 15, EVF_used
  EditBox 60, 165, 320, 15, Reason_OP
  CheckBox 60, 190, 120, 10, "Earned Income disregard allowed", EI_checkbox
  ButtonGroup ButtonPressed
    OkButton 285, 185, 45, 15
    CancelButton 335, 185, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 95, 10, 30, 10, "MEMB #:"
  Text 160, 10, 60, 10, "Other resp. memb:"
  Text 255, 10, 55, 10, "Discovery Date: "
  GroupBox 5, 25, 385, 85, "Overpayment Information"
  Text 10, 50, 30, 10, "Program:"
  Text 100, 50, 20, 10, "From:"
  Text 180, 50, 10, 10, "To:"
  Text 245, 50, 25, 10, "Claim #"
  Text 320, 50, 20, 10, "AMT:"
  Text 10, 70, 30, 10, "Program:"
  Text 100, 70, 20, 10, "From:"
  Text 180, 70, 10, 10, "To:"
  Text 245, 70, 25, 10, "Claim #"
  Text 320, 70, 20, 10, "AMT:"
  Text 10, 90, 30, 10, "Program:"
  Text 100, 90, 20, 10, "From:"
  Text 180, 90, 10, 10, "To:"
  Text 245, 90, 25, 10, "Claim #"
  Text 320, 90, 20, 10, "AMT:"
  Text 5, 125, 40, 10, "Collectible?"
  Text 95, 125, 65, 10, "Collectible Reason:"
  Text 290, 125, 50, 10, "Fraud referral:"
  Text 5, 145, 50, 10, "Income Source: "
  Text 150, 145, 85, 10, "Income verification used:"
  Text 5, 170, 50, 10, "Reason for OP: "
  Text 125, 35, 15, 10, "(MM)"
  Text 155, 35, 15, 10, "(YY)"
  Text 195, 35, 15, 10, "(MM)"
  Text 225, 35, 15, 10, "(YY)"
EndDialog



Do
	err_msg = ""
	dialog EWS_OP_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbnewline & "* Enter a valid case number."
IF First_Program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment FS Food Stamps, FG Family GA, GA Gen Assist, GR Group Residential Housing, MF MFIP, or DW Diversionary Work Program"
	IF First_OP = "" THEN err_msg = err_msg & vbnewline & "* You must have an overpayment entry."
	IF First_from_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start month(MM) overpayment occured."
	IF First_from_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start year(YY) overpayment occured."
	IF First_to_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end month(MM) overpayment occured."
	IF First_to_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end year(YY) overpayment occured."
	IF Second_OP <> "" THEN
		IF Second_from_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month(MM) 2nd overpayment occured."
		IF Second_from_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start year(YY) 2nd overpayment occured."
		IF Second_to_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end month(MM) 2nd overpayment occured."
		IF Second_to_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end year(YY) 2nd overpayment occured."
	END IF
	IF Third_OP <> "" THEN
		IF Third_from_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month(MM) 3rd overpayment occured."
		IF Third_from_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start year(YY) 3rd overpayment occured."
		IF Third_to_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end month(MM) 3rd overpayment occured."
		IF Third_to_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end year(YY) 3rd overpayment occured."
	END IF
	IF collectible_dropdown = "Select:"  THEN err_msg = err_msg & vbnewline & "* Please advise if overpayment is collectible."
	IF collectible_dropdown = "YES"  & collectible_reason = "" THEN err_msg = err_msg & vbnewline & "* Please advise why overpayment is collectible."
	IF fraud_referral = "Select:"  THEN err_msg = err_msg & vbnewline & "* Please advise if a fraud referral was made."
	IF source_income = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the source of income."
	IF EVF_used = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the verifcation used for income."
	IF Reason_OP = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the reason for overpayment."
	IF Discovery_date = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the date the overpayment was discovered (DD/MM/YY)."
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
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("----- " & First_from_IEVS_month & "/" &  First_from_IEVS_year & "(" & first_name &  ")" & "OVERPAYMENT CLAIM ENTERED -----")
CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ")
Call write_variable_in_CASE_NOTE(First_Program & " Overpayment Claim # " & First_OP &  " " & First_from_IEVS_month & "/" &  First_from_IEVS_year & " through "  & First_to_IEVS_month & "/" &  First_to_IEVS_year & " Amount: " & First_amount)
IF Second_OP <> "" THEN CALL write_variable_in_case_note(Second_Program &  "Overpayment Claim # " & Second_OP &  " " & Second_from_IEVS_month & "/" &  Second_from_IEVS_year & " through "  & Second_to_IEVS_month & "/" &  Second_to_IEVS_year & " Amount: " & Second_amount)
IF Third_OP <> "" THEN CALL write_variable_in_case_note(Third_Program &  "Overpayment Claim # " & Third_OP &  " " & Third_from_IEVS_month & "/" &  Third_from_IEVS_year & " through "  & Third_to_IEVS_month & "/" &  Third_to_IEVS_year & " Amount: " & Third_amount)
IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
IF fraud_referral = "YES" THEN CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
CALL write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
CALL write_bullet_and_variable_in_case_note("Verification used for overpayment", EVF_used)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_case_note("Discovery Date", Discovery_date)
CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
CALL write_variable_in_CASE_NOTE("----- ----- -----")
CALL write_variable_in_CASE_NOTE(worker_signature)

PF3

script_end_procedure("Overpayment case note entered. Please remember to copy and paste your notes to CCOL/CLIC")
