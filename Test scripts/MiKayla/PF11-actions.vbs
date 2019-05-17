'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PF11 ACTIONS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 120                	'manual run time in seconds
STATS_denomination = "C"       		'M is for each MEMBER
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
call changelog_update("05/13/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'intial dialog for user to select a SMRT action
BeginDialog , 0, 0, 186, 65, "SMRT initial dialog"
  EditBox 85, 5, 60, 15, maxis_case_number
  DropListBox 85, 25, 95, 15, "Select one..."+chr(9)+"Initial request"+chr(9)+"ISDS referral completed"+chr(9)+"Determination received", SMRT_actions
  ButtonGroup ButtonPressed
    OkButton 75, 45, 50, 15
    CancelButton 130, 45, 50, 15
  Text 5, 30, 75, 10, "Select a SMRT action:"
  Text 30, 10, 45, 10, "Case number:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog
		if ButtonPressed = 0 then StopScript
		if IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If SMRT_actions = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select a SMRT action."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)

'Error proof functions
Call check_for_MAXIS(true)

DO
	dialog EDRS_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MSGBOX "Please enter a case number"

LOOP UNTIL MAXIS_case_number <> ""

'Error proof functions
Call check_for_MAXIS(False)

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Error proof functions
Call check_for_MAXIS(False)

'changing footer dates to current month to avoid invalid months.
MAXIS_footer_month = datepart("M", date)
	IF Len(MAXIS_footer_month) <> 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", date), 2)

Dim Member_Info_Array()
Redim Member_Info_Array(UBound(HH_member_array), 4)


'Navigate to stat/memb and check for ERRR message
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
For i = 0 to Ubound(HH_member_array)

	Member_Info_Array(i, 0) = HH_member_array(i)
	'Navigating to selected memb panel
	EMwritescreen HH_member_array(i), 20, 76
	transmit

	EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
	IF no_MEMB = "Arrival Date:" THEN script_end_procedure("This HH member does not exist.")


	'Reading info and removing spaces
	EMReadscreen First_name, 12, 6, 63
	First_name = replace(First_name, "_", "")
	Member_Info_Array(i, 1) = First_name

	'Reading Last name and removing spaces
	EMReadscreen Last_name, 25, 6, 30
	Last_name = replace(Last_name, "_", "")
	Member_Info_Array(i, 2) = Last_name

	'Reading Middle initial and replacing _ with a blank if empty.
	EMReadscreen Middle_initial, 1, 6, 79
	Middle_initial = replace(Middle_initial, "_", "")
	Member_Info_Array(i, 3) = Middle_initial

	'Reads SSN
	Emreadscreen SSN_number, 11, 7, 42
	SSN_number = replace(SSN_number, " ", "")
	Member_Info_Array(i, 4) = SSN_number
Next

BeginDialog , 0, 0, 361, 75, "PF11 Action"
  EditBox 55, 20, 40, 15, maxis_case_number
  DropListBox 175, 20, 95, 15, "Select one:"+chr(9)+"PMI merge request"+chr(9)+"Non-actionable DAIL removal"+chr(9)+"Case note removal request", PF11_actions
  ButtonGroup ButtonPressed
    OkButton 250, 55, 50, 15
    CancelButton 305, 55, 50, 15
  Text 110, 25, 65, 10, "Select PF11 action:"
  Text 5, 25, 45, 10, "Case number:"
  Text 5, 5, 355, 10, "The system being down, issuance problems, or any type of emergency SHOULD NOT be reported via a PF11."
  Text 30, 40, 245, 10, "MAXIS takes a "snapshot" of the screen you choose to do your PF11 from.  "
  EditBox 335, 20, 20, 15, MEMB_number
  Text 280, 25, 50, 10, "MEMB number:"
EndDialog


If SMRT_actions = "Initial request" then
    BeginDialog , 0, 0, 326, 180, "Initial SMRT referral dialog"
      EditBox 80, 10, 75, 15, SMRT_member
      EditBox 270, 10, 50, 15, referral_date
      DropListBox 80, 35, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", referred_exp
      EditBox 200, 35, 120, 15, expedited_reason
      EditBox 80, 60, 240, 15, referral_reason
      EditBox 80, 85, 50, 15, SMRT_start_date
      If worker_county_code = "x127" then CheckBox 140, 90, 180, 10, "Check here if the ECF workflow has been completed.", ECF_workflow_checkbox
      EditBox 80, 110, 240, 15, other_notes
      EditBox 80, 135, 240, 15, action_taken
      EditBox 80, 160, 130, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 215, 160, 50, 15
        CancelButton 270, 160, 50, 15
      Text 15, 115, 65, 10, "Other SMRT notes:"
      Text 5, 40, 70, 10, "Is referral expedited?"
      Text 25, 140, 50, 10, " Actions taken:"
      Text 165, 15, 100, 10, "Date SMRT referral completed:"
      Text 5, 15, 70, 10, "SMRT requested for: "
      Text 20, 90, 60, 10, "SMRT start date:"
      Text 15, 165, 60, 10, "Worker Signature:"
      Text 155, 40, 45, 10, "If yes, why?:"
      Text 10, 65, 65, 10, "Reason for referral:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog
    		cancel_confirmation
    		If SMRT_member = "" THEN err_msg = err_msg & vbNewLine & "* Enter the member info the SMRT referral."
    		If isdate(referral_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid referral date."
			If referred_exp = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Is the referral expedited?"
			If (referred_exp = "Yes" and trim(expedited_reason) = "") THEN err_msg = err_msg & vbNewLine & "* Enter the expedited reason."
			If trim(referral_reason) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the reason for the referral."
			If isdate(SMRT_start_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid SMRT start date."
			If trim(action_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
			If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False

	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("---Initial SMRT referral requested---")
	call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
	Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral completed on", referral_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Is referral expedited", referred_exp)
	If referred_exp = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Expedited reason", expedited_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("SMRT start date", SMRT_start_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	If ECF_workflow_checkbox = 1 then call write_variable_in_CASE_NOTE("* ECF workflow has been completed in ECF.")
	Call write_variable_in_CASE_NOTE ("---")
	call write_variable_in_CASE_NOTE(worker_signature)
END If
Please delete DAIL - case is inactive and cant be resolved


***PF11 SENT RE: DAIL REMOVAL***
TASK #438668
----
Could not remove PARI DAIL, case inactive, PF11 sent for removal.
---
M. Handley Quality Improvement

Describe the problem:
QTIP #108 - PF11 BASICS                       TE19.108             3 of
In the DESCRIBE PROBLEM field:

   >  Explain exactly what you are trying to do.

   >  Explain what MAXIS is doing and what results you are
   expecting.

>  Enter your direct dial telephone number.
expecting.
If the screen you have chosen is DAIL/DAIL, REPT/PND2, or any
screen that displays a list of cases please identify the case
number you are having problems with as it will not automatically
display in the CASE ID field.

When you have finished describing the problem and wish to complete
the PF11, transmit.  If for some reason you have decided not to
send the PF11 hit the PF11 key to cancel.

Once you have completed the PF11 a task number will appear.  Keep
the task number for future reference.

To check on a PF11's progress go to the SELF menu and type TASK.
If you have the task number enter it and it will take you directly
QTIP #108 - PF11 BASICS                       TE19.108             3 of
into the PF11.  If you do not have the task number or wish to look
at a list of all the PF11s you have created, change the Option in
TASK from "T" (task) to a "C" (creator).  By placing an "X" next to
a PF11 listed you will be able to view it.

PF11s are printed twice a day:  7:00 a.m. and noon.  They are
reviewed as soon as possible and distributed to the appropriate
group for action.

For more information on PF11s see the TEMP Manual TE05.02 (TSS Help
Desk Procedures), TE05.04 (PF11's), and TE19.037 (QTIP #37 - PF11
Assignment Process).

If you have questions about this QTIP or suggestions for future
QTIPS please e-mail QTIP.



please delete duplicate case note

script_end_procedure("It worked!")
