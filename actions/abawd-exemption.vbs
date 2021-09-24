'GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - ABAWD EXEMPTION.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
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
call changelog_update("05/23/2018", "Bug fix for living situation coding inhibiting users from using code 08.", "Ilse Ferris, Hennepin County")
call changelog_update("04/17/2018", "Added inhibiting coding for homeless (Unfit for Employement) if the ADDR panel is not coded correctly.", "Ilse Ferris, Hennepin County")
call changelog_update("03/29/2018", "Added Homeless (Unfit for Employment) option.", "Ilse Ferris, Hennepin County")
call changelog_update("09/07/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & case number
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
member_number = "01"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 90, "ABAWD exemption dialog"
  EditBox 70, 10, 50, 15, MAXIS_case_number
  EditBox 170, 10, 20, 15, member_number
  EditBox 135, 30, 55, 15, effective_date
  DropListBox 70, 50, 120, 15, "Select one..."+chr(9)+"Care of Child under 6"+chr(9)+"Care of Incapacitated Person"+chr(9)+"Homeless (Unfit for Employment)"+chr(9)+"Other", ABAWD_selection
  ButtonGroup ButtonPressed
    OkButton 90, 70, 50, 15
    CancelButton 140, 70, 50, 15
  Text 40, 35, 90, 10, "Effective date of exemption:"
  Text 20, 15, 45, 10, "Case number:"
  Text 130, 15, 35, 10, "Member #:"
  Text 5, 55, 65, 10, "ABAWD exemption:"
EndDialog
'the dialog
Do
	Do
  		err_msg = ""
  		Dialog Dialog1
  		If ButtonPressed = 0 then stopscript
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid member number."
		If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
		If ABAWD_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an ABWAD exemption."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If ABAWD_selection = "Homeless (Unfit for Employment)" then
    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, living_situation, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
    If addr_homeless <> "Yes" then script_end_procedure("This case does not have the ADDR panel coded as homeless. Please review the case, and run the script again as needed.")
	living_situation = left(living_situation, 2)
	If  living_situation = "01" or _
        living_situation = "03" or _
        living_situation = "04" or _
        living_situation = "05" or _
        living_situation = "09" or _
        living_situation = "10" or _
        living_situation = "Bl" then
        script_end_procedure("This case's living situation code on the ADDR panel does not meet this exemption criteria. Please review the case, and run the script again as needed.")
    Else
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 281, 170, "Unfit for Employment exemption for homeless members"
          EditBox 215, 90, 60, 15, conversation_date
          EditBox 75, 110, 200, 15, exemption_details
          EditBox 75, 130, 200, 15, other_notes
          EditBox 75, 150, 90, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 170, 150, 50, 15
            CancelButton 225, 150, 50, 15
          Text 15, 155, 60, 10, "Worker signature:"
          GroupBox 5, 10, 270, 75, "Clients will meet this exemption if BOTH criteria are met:"
          Text 10, 25, 260, 25, "* The client is coded Y as homeless on STAT/ADDR with a living arrangement code of either (02) Family/Friends Due to Economic Hardship, (06) Hotel/Motel, (07) Emergency Shelter,or (08) Place Not Meant for Housing AND"
          Text 10, 60, 260, 20, "* The client lacks access to work-related necessities. These necessities include, but are not limited to, access to a shower and/or laundry facilities."
          Text 45, 95, 170, 10, "Conversation date with client about this exemption:"
          Text 30, 135, 45, 10, "Other notes:"
          Text 10, 115, 65, 10, "Exemption details:"
        EndDialog

        'the Other exemption dialog
        Do
            Do
                err_msg = ""
                Dialog Dialog1
                Cancel_confirmation
                If isdate(conversation_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid conversation date."
                If trim(exemption_details) = "" then err_msg = err_msg & vbNewLine & "* Enter the details of your conversation with the client with the specifics about why they met this exemption."
                If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
            LOOP UNTIL err_msg = ""
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in
    End if
elseIf ABAWD_selection<> "Other" then
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 281, 190, "ABAWD exemption for " & ABAWD_selection
  	  EditBox 40, 25, 155, 15, person_name
  	  If ABAWD_selection = "Care of Child under 6" then EditBox 225, 25, 40, 15, person_DOB
  	  EditBox 130, 45, 20, 15, hours_per_week
  	  EditBox 245, 45, 20, 15, days_per_week
  	  CheckBox 15, 65, 250, 10, "Check if more than one HH memb is claiming an exemption for their care.", extra_person_checkbox
  	  DropListBox 85, 105, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", verifs_required
  	  DropListBox 215, 105, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", verifs_rec
  	  EditBox 65, 125, 200, 15, verif_info
  	  EditBox 65, 150, 200, 15, other_notes
  	  EditBox 65, 170, 90, 15, worker_signature
  	  ButtonGroup ButtonPressed
	  OkButton 160, 170, 50, 15
	  CancelButton 215, 170, 50, 15
	  If ABAWD_selection = "Care of Child under 6" then Text 205, 30, 20, 10, "DOB:"
  	  Text 15, 30, 25, 10, "Name:"
  	  Text 15, 50, 110, 10, "How many care hours per week?"
  	  Text 15, 155, 45, 10, "Other notes:"
  	  Text 170, 50, 70, 10, "Care days per week?"
  	  GroupBox 5, 90, 265, 55, "Verification is required if info is questionable/more than 1 exemption claimed"
  	  Text 5, 175, 60, 10, "Worker signature:"
  	  GroupBox 5, 10, 265, 70, "The following fields are information about the person in the ABWAD's care"
  	  Text 145, 110, 70, 10, "Verification received:"
  	  Text 15, 110, 70, 10, "Verification required:"
  	  Text 10, 130, 55, 10, "Verification info:"
	EndDialog

	'the Incap/child under 6 dialog
	Do
		Do
	  		err_msg = ""
	  		Dialog Dialog1
	  		Cancel_confirmation
			If trim(person_name) = "" then err_msg = err_msg & vbNewLine & "* Enter the person in care's first and last name."
			If (ABAWD_selection = "Care of Child under 6" AND IsDate(person_DOB) = False) then err_msg = err_msg & vbNewLine & "* Enter a valid DOB for child under 6."
			IF IsNumeric(hours_per_week) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric hours per week."
	  		IF IsNumeric(days_per_week) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric days per week."
			IF hours_per_week < 18.60 then err_msg = err_msg & vbNewLine & "*ABAWD exemption can only be used equal to or greater than 20 hours per/week or 80 hours/month. Press CANCEL if client doesn't meet this criteria."
			If verifs_required = "Select one..." then err_msg = err_msg & vbNewLine & "* Were verifications required for this exemption?"
			If (verifs_required = "Yes" AND verifs_rec = "Select one...") then err_msg = err_msg & vbNewLine & "* Were verifications received for this exemption?"
			If (extra_person_checkbox = 1 AND verifs_required = "No") then err_msg = err_msg & vbNewLine & "* Verifications must be provided for exemption if more than 1 HH member is claiming the same person for their exemption."
			If (verifs_required = "Yes" AND verifs_rec = "Yes" AND trim(verif_info) = "") then err_msg = err_msg & vbNewLine & "* Complete the 'verification info' field."
			If (verifs_required = "Yes" AND verifs_rec = "No" AND trim(verif_info) = "") then err_msg = err_msg & vbNewLine & "* Explain why client is eligible for exemption without verifications received in the 'verification info' field."
			If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
Else
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 276, 140, "ABAWD exemption: Select first code available"
	  'This droplist is too damn big to enter into the dialog editor. You WILL break the dialog editor if you paste this code into it.
	  DropListBox 75, 10, 190, 15, "Select one..."+chr(9)+"03 Unfit for Employment"+chr(9)+"05 Age 60 or older"+chr(9)+"06 Under age 16"+chr(9)+"07 Age 16-17 living w/ parent/caregiver"+chr(9)+"09 Empl 30 hr/wk or earnings = to min wage x 30 hr/wk"+chr(9)+"10 Matching grant participant"+chr(9)+"11 Receiving or applied for unemployment"+chr(9)+"12 Enrolled in school, training program or higher education"+chr(9)+"13 Participating In CD Program"+chr(9)+"14 Receiving MFIP"+chr(9)+"20 Pending/Receiving DWP Or WB"+chr(9)+"15 Age 16-17 Not Lvg W/Pare/Crgvr"+chr(9)+"16 50-59 years old"+chr(9)+"21 Resp For Care Of Child < 18"+chr(9)+"17 Receiving RCA Or GA", Exemption_droplist
  	  DropListBox 80, 50, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", verifs_required
  	  DropListBox 215, 50, 50, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", verifs_rec
  	  EditBox 65, 70, 200, 15, verif_info
  	  EditBox 65, 100, 200, 15, other_notes
  	  EditBox 65, 120, 90, 15, worker_signature
  	  ButtonGroup ButtonPressed
    	OkButton 160, 120, 50, 15
    	CancelButton 215, 120, 50, 15
  	  Text 5, 125, 60, 10, "Worker signature:"
  	  Text 140, 55, 70, 10, "Verification received:"
  	  Text 10, 55, 70, 10, "Verification required:"
  	  Text 10, 75, 55, 10, "Verification info:"
  	  Text 10, 15, 65, 10, "ABAWD exemption:"
  	  Text 20, 105, 45, 10, "Other notes:"
  	  GroupBox 5, 35, 265, 55, "Verification of exemption"
	EndDialog

	'the Other exemption dialog
	Do
		Do
	  		err_msg = ""
	  		Dialog Dialog1
	  		Cancel_confirmation
			If Exemption_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an ABAWD exemption."
			If verifs_required = "Select one..." then err_msg = err_msg & vbNewLine & "* Were verifications required for this exemption?"
			If (verifs_required = "Yes" AND verifs_rec = "Select one...") then err_msg = err_msg & vbNewLine & "* Were verifications received for this exemption?"
			If (verifs_required = "Yes" AND verifs_rec = "Yes" AND trim(verif_info) = "") then err_msg = err_msg & vbNewLine & "* Complete the 'verification info' field."
			If (verifs_required = "Yes" AND verifs_rec = "No" AND trim(verif_info) = "") then err_msg = err_msg & vbNewLine & "* Explain why client is eligible for exemption without verifications received in the 'verification info' field."
			If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
End if

MAXIS_footer_month 	= right("0" & DatePart("m",   effective_date), 2)
MAXIS_footer_year 	= right(      DatePart("yyyy",effective_date), 2)

MAXIS_footer_month_confirmation

Call navigate_to_MAXIS_screen("STAT", "WREG")
Do
	EMReadScreen WREG_panel, 4, 2, 48
	If WREG_panel <> "WREG" then Call navigate_to_MAXIS_screen("STAT", "WREG")
Loop until WREG_panel = "WREG"

CALL write_value_and_transmit(member_number, 20, 76)
EMReadScreen WREG_MEMB_check, 6, 24, 2
IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER" THEN script_end_procedure ("The member number that you entered is not valid.  Please check the member number, and start the script again.")

EMReadscreen wreg_panel, 1, 2, 78
If wreg_panel = "0" then Call write_value_and_transmit("NN", 20, 79)
EMReadscreen PWE_indicator, 1, 6, 68
If PWE_indicator = "_" then EMWriteScreen "Y", 6, 68

If ABAWD_selection = "Care of Child under 6" then
	FSET_exemption_code = "08"
	ABAWD_input_code = "01"
Elseif ABAWD_selection = "Care of Incapacitated Person" then
	FSET_exemption_code = "04"
	ABAWD_input_code = "01"
Elseif ABAWD_selection = "Homeless (Unfit for Employment)" then
	FSET_exemption_code = "03"
	ABAWD_input_code = "01"
Else
	'Logic to change the GA_basis_droplist into correct coding for the WREG panel
	FSET_exemption_code = Left(Exemption_droplist, 2)
	'Determining what the ABAWD code will be based on the FSET code (per POLI TEMP)
	If FSET_exemption_code = "03" then ABAWD_input_code = "01"
	If FSET_exemption_code = "05" then ABAWD_input_code = "01"
	If FSET_exemption_code = "06" then ABAWD_input_code = "01"
	If FSET_exemption_code = "07" then ABAWD_input_code = "01"
	If FSET_exemption_code = "09" then ABAWD_input_code = "01"
	If FSET_exemption_code = "10" then ABAWD_input_code = "01"
	If FSET_exemption_code = "11" then ABAWD_input_code = "01"
	If FSET_exemption_code = "12" then ABAWD_input_code = "01"
	If FSET_exemption_code = "13" then ABAWD_input_code = "01"
	If FSET_exemption_code = "14" then ABAWD_input_code = "01"
	If FSET_exemption_code = "20" then ABAWD_input_code = "01"

	If FSET_exemption_code = "15" then ABAWD_input_code = "02"
	If FSET_exemption_code = "16" then ABAWD_input_code = "03"
	If FSET_exemption_code = "17" then ABAWD_input_code = "12"
	If FSET_exemption_code = "21" then ABAWD_input_code = "04"
End if

EMReadScreen FSET_code, 2, 8, 50
EMReadScreen ABAWD_code, 2, 13, 50

'Detemining if WREG panel will need to be updated or not
If FSET_code = FSET_exemption_code then
	If ABAWD_code = ABAWD_input_code then update_WREG = False
Else
	update_WREG = True
End if

'script will update the WREG panel for the member if an update
If update_WREG = true then
	PF9
	EMWriteScreen FSET_exemption_code, 8, 50
	EMWriteScreen ABAWD_input_code, 13, 50
	EMWriteScreen "_", 8, 80
	PF3
End if

If ABAWD_selection = "Other" then
	header_info = right(Exemption_droplist, len(Exemption_droplist) - 3)
Else
	header_info = ABAWD_selection
End if

'----------------------------------------------------------------------------------------------------The case note
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("ABAWD Exemption for M" & member_number & ": " & header_info)
Call write_bullet_and_variable_in_CASE_NOTE("Effective date", effective_date)
Call write_bullet_and_variable_in_CASE_NOTE("Date of conversation with client", conversation_date)
Call write_bullet_and_variable_in_CASE_NOTE("Why client meets Unfit to Employment expansion for homelessness", exemption_details)
Call write_bullet_and_variable_in_CASE_NOTE("Person in care of ABAWD", person_name)
Call write_bullet_and_variable_in_CASE_NOTE("DOB of person in care of ABAWD", person_DOB)
Call write_bullet_and_variable_in_CASE_NOTE("Hours a week caring for person", hours_per_week)
Call write_bullet_and_variable_in_CASE_NOTE("Days a week caring for person", days_per_week)
Call write_bullet_and_variable_in_CASE_NOTE("Verifications required", verifs_required)
If verifs_rec <> "Select one..." then Call write_bullet_and_variable_in_CASE_NOTE("Verification received", verifs_rec)
Call write_bullet_and_variable_in_CASE_NOTE("Verification info", verif_info)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(Worker_Signature)
If update_WREG = true then msgbox "The WREG panel has been updated to reflect this exemption for the effective footer month/year." & vbcr & "Please review the member's ABAWD tracking record and delete any months the member meets an exemption."

script_end_procedure("")
