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
'FUNCTIONS==================================================================================================================
Function Generate_Client_List(list_for_dropdown)

	memb_row = 5

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		transmit
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20

	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~")

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next

End Function
'THE SCRIPT=================================================================================================================
Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


If MAXIS_case_number <> "" Then 		'If a case number is found the script will get the list of
	Call Generate_Client_List(HH_Memb_DropDown)
End If

'Running the dialog for case number and client
Do
	err_msg = ""
	'Dialog defined here so the dropdown can be changed
	BeginDialog select_person_dialog, 0, 0, 191, 65, "Select Caregiver"
	  EditBox 55, 5, 50, 15, MAXIS_case_number
	  ButtonGroup ButtonPressed
	    PushButton 135, 5, 50, 15, "search", search_button
	  DropListBox 80, 25, 105, 45, "Select One..." & HH_Memb_DropDown, clt_to_update
	  ButtonGroup ButtonPressed
	    OkButton 115, 45, 35, 15
	    CancelButton 155, 45, 30, 15
	  Text 5, 10, 45, 10, "Case Number"
	  Text 5, 30, 70, 10, "Household member"
	EndDialog

	Dialog select_person_dialog
	If ButtonPressed = cancel Then StopScript
	If ButtonPressed = search_button Then
		If MAXIS_case_number = "" Then
			MsgBox "Cannot search without a case number, please try again."
		Else
			HH_Memb_DropDown = ""
			Call Generate_Client_List(HH_Memb_DropDown)
			err_msg = err_msg & "Start Over"
		End If
	End If
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a valid case number."
	If clt_to_update = "Select One..." Then err_msg = err_msg & vbNewLine & "Please pick a client whose EMPS panel you need to update."
	If err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""



'intial dialog for user to select a SMRT action
BeginDialog , 0, 0, 361, 80, "PF11 Action"
  EditBox 70, 20, 40, 15, maxis_case_number
  EditBox 70, 40, 15, 15, MEMB_number
  DropListBox 205, 40, 95, 15, "Select One:"+chr(9)+"PMI merge request"+chr(9)+"Non-actionable DAIL removal"+chr(9)+"Case note removal request"+chr(9)+"MFIP New Spouse Income"+chr(9)+"Other", PF11_actions
  EditBox 70, 60, 170, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 60, 50, 15
    CancelButton 305, 60, 50, 15
  Text 5, 5, 355, 10, "The system being down, issuance problems, or any type of emergency SHOULD NOT be reported via a PF11."
  Text 115, 25, 245, 10, "** MAXIS takes a snapshot of the screen you choose to do your PF11 from."
  Text 135, 45, 65, 10, "Select PF11 action:"
  Text 5, 45, 50, 10, "MEMB number:"
  Text 5, 65, 60, 10, "Worker Signature:"
  Text 5, 25, 45, 10, "Case number:"
EndDialog


BeginDialog , 0, 0, 326, 90, "Other"
  EditBox 70, 10, 240, 15, referral_reason
  EditBox 70, 30, 240, 15, other_notes
  EditBox 70, 50, 240, 15, action_taken
  ButtonGroup ButtonPressed
    OkButton 205, 70, 50, 15
    CancelButton 260, 70, 50, 15
  Text 25, 35, 45, 10, "Other Notes:"
  Text 15, 55, 50, 10, " Actions Taken:"
  Text 5, 15, 60, 10, "Describe Problem:"
EndDialog


Do
	Do
		err_msg = ""
		Dialog
		if ButtonPressed = 0 then StopScript
		if IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If SMRT_actions = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Select a PF11 action."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

If PF11_actions = "PMI merge request" then
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

script_end_procedure("It worked!")
