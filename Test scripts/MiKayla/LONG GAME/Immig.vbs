'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTION-IMMIGRATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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
call changelog_update("03/28/2018", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG PORTION----------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog IMIG_dialog, 0, 0, 366, 280, "Immigration"
  EditBox 65, 5, 40, 15, MAXIS_case_number
	EditBox 175, 5, 20, 15, memb_number
	EditBox 275, 5, 70, 15, alien_id_number
  CheckBox 260, 10, 85, 10, "Inital SAVE requested?", save_requested_check
	DropListBox 60, 35, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More "+chr(9)+"26 Conditional Entry < 4/80 "+chr(9)+"27 Non-Immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing", Immig_status_dropdown
  DropListBox 60, 55, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More "+chr(9)+"26 Conditional Entry < 4/80 "+chr(9)+"27 Non-Immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"N/A", LPR_status_dropdown
  DropListBox 255, 75, 95, 15, "Select One:"+chr(9)+"Certificate of Naturalization"+chr(9)+"Employment Auth Card (I-776 work permit)"+chr(9)+"I-94 Travel Document", +chr(9)+"I-220 B Order of Supervision"+chr(9)+"LPR Card (I-551 green card)"+chr(9)+"SAVE"+chr(9)+"Other", immig_doc_type
  DropListBox 255, 55, 95, 15, "Select One:"+chr(9)+"AA Amerasian"+chr(9)+"EH Ethnic Chinese"+chr(9)+"EL Ethnic Lao"+chr(9)+"HG Hmong"+chr(9)+"KD Kurd"+chr(9)+"SJ Soviet Jew"+chr(9)+"TT Tinh"+chr(9)+"AF Afghanistan"+chr(9)+"BK Bosnia"+chr(9)+"CB Cambodia"+chr(9)+"CH China,"+chr(9)+"Mainland"+chr(9)+"CU Cuba"+chr(9)+"ES El Salvador"+chr(9)+"ER Eritrea"+chr(9)+"ET Ethiopia"+chr(9)+"GT Guatemala"+chr(9)+"HA Haiti "+chr(9)+"HO Honduras"+chr(9)+"IR Iran"+chr(9)+"IZ Iraq"+chr(9)+"LI Liberia"+chr(9)+"MC Micronesia"+chr(9)+"MI Marshall"+chr(9)+"Islands"+chr(9)+"MX Mexico"+chr(9)+"WA Namibia"+chr(9)+"(SW Africa)"+chr(9)+"PK Pakistan"+chr(9)+"RP Philippines"+chr(9)+"PL Poland"+chr(9)+"RO Romania"+chr(9)+"RS", nationality_dropdown
  DropListBox 255, 35, 95, 15, "Select One:"+chr(9)+"SAVE Primary"+chr(9)+"SAVE Secondary"+chr(9)+"AL Alien Card"+chr(9)+"PV Passport/Visa "+chr(9)+"RE Re-Entry Prmt "+chr(9)+"IM INS Correspondence"+chr(9)+"OT Other Document"+chr(9)+"NO No Ver Prvd ", status_verification
	EditBox 60, 75, 45, 15, date_of_entry
	CheckBox 10, 105, 85, 10, "Inital SAVE requested?", save_requested_check
  CheckBox 120, 105, 100, 10, "Additional SAVE requested?", additional_save_check
  CheckBox 10, 120, 215, 10, "If checked did you attach a copy of the immigration document?", SAVE_docs_check
  OptionGroup RadioGroup1
    RadioButton 15, 155, 25, 10, "No", not_sponsored
    RadioButton 15, 170, 75, 10, "Yes, sponsored by:", sponsored
  EditBox 85, 190, 70, 15, name_sponsor
  EditBox 220, 190, 100, 15, sponsor_addr
  EditBox 85, 210, 70, 15, name_sponsor_two
  EditBox 220, 210, 100, 15, sponsor_addr_two
  EditBox 85, 230, 70, 15, name_sponsor_three
  EditBox 220, 230, 100, 15, sponsor_addr_three
  EditBox 55, 255, 135, 15, other_notes
	ButtonGroup ButtonPressed
		OkButton 260, 255, 45, 15
		CancelButton 310, 255, 45, 15
  Text 10, 10, 50, 10, "Case Number:"
	Text 115, 10, 60, 10, "Member Number:"
	Text 230, 10, 40, 10, "Alien ID # A:"
	GroupBox 5, 25, 350, 110, "Immigration Information"
  Text 10, 40, 50, 10, "Immig. Status:"
  Text 10, 55, 45, 15, "LPR adjusted from:"
  Text 200, 40, 50, 10, "Status Verified:"
  Text 10, 80, 50, 10, "Date of entry:"
  Text 190, 60, 60, 10, "Nationality/Nation:"
  Text 200, 80, 55, 10, "Immig doc type:"
  GroupBox 5, 140, 350, 110, "Sponsored on I-864 Affidavit of Support? (LPR COA CODE:  C, CF, CR, CX, F, FX, IF, IR)"
  Text 105, 155, 240, 10, "*If date of entry was prior to 12/19/1997 sponsor information is not needed"
  Text 120, 170, 205, 10, "*If sponsor is active on MAXIS case information is not needed"
  Text 20, 195, 60, 10, "Name of sponsor:"
  Text 165, 195, 55, 10, "Address/Phone:"
  Text 20, 215, 60, 10, "Name of sponsor:"
  Text 165, 215, 55, 10, "Address/Phone:"
  Text 20, 235, 60, 10, "Name of sponsor:"
  Text 165, 235, 55, 10, "Address/Phone:"
  Text 10, 260, 45, 10, "Other Notes:"
EndDialog


'THE SCRIPT PORTION----------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)      'finding case number
Call check_for_MAXIS(true)						'making sure that person is in MAXIS and logged in

Do
	Do
		err_msg = ""
		Dialog IMIG_dialog
		cancel_confirmation
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF memb_number = "" or IsNumeric(memb_number) = False or len(memb_number) > 2 then err_msg = err_msg & vbNewLine & "* Enter a member number."
		IF alien_id_number = "" or IsNumeric(alien_id_number) = False or len(alien_id_number) <> 9  then err_msg = err_msg & vbNewLine & "* Enter immigration ID number, must be 9 digits and numeric only."
		IF save_requested_check = UNCHECKED then err_msg = err_msg & vbNewLine & "* Please select if a SAVE has been run as it is mandatory."
		IF Immig_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of current immigration status."
		IF Immig_status_dropdown = "24 LPR" and LPR_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of LPR adjusted status."
		IF immig_doc_type= "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of immigration document used."
		IF nationality_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of Nationality or Nation."
		IF sponsored = 1 and name_sponsor = "" then err_msg = err_msg & vbNewLine & "* You indicated a sponsor for this case please complete sponsor information."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen("STAT", "IMIG")
'Making sure we have the correct IMIG
EMReadScreen panel_number, 1, 2, 78
If panel_number = "0" then script_end_procedure("An IMIG panel does not exist. Please create the panel before running the script again. ")
'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
Do
	EMReadScreen current_panel_number, 1, 2, 73
	IMIG_check = MsgBox("Is this the right IMIG?", vbYesNo +vbQuestion, "Confirmation")
	If IMIG_check = vbYes then exit do
	If IMIG_check = vbNo then	TRANSMIT
	If (IMIG_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure("Unable to find another IMIG. Please review the case, and run the script again if applicable.")
	Loop until current_panel_number = panel_number
End if

'Updating the IMIG panel
PF9
EMReadScreen error_check, 2, 24, 2	'making sure we can actually update this case.
error_check = trim(error_check)
If error_check <> "" then script_end_procedure("Unable to update this case. Please review case, and run the script again if applicable.")

EMWriteScreen "Y", 4, 73			'Support Coop Y/N field
EMWriteScreen "P", 5, 47			'Good Cause status field
EMWriteScreen "N", 7, 47			'Sup evidence Y/N field (defaulted to N during this process)
Call create_MAXIS_friendly_date(claim_date, 0, 5, 73)

'converting the good cause reason from reason_droplist to the applicable MAXIS coding
If reason_droplist = "Potential phys harm/Child"		then claim_reason = "1"
If reason_droplist = "Potential Emotnl harm/Child"	 	then claim_reason = "2"
If reason_droplist = "Potential phys harm/Caregiver" 	then claim_reason = "3"
If reason_droplist = "Potential Emotnl harm/Caregiver" 	then claim_reason = "4"
If reason_droplist = "Cncptn Incest/Forced Rape" 		then claim_reason = "5"
If reason_droplist = "Legal adoption Before Court" 		then claim_reason = "6"
If reason_droplist = "Parent Gets Preadoptn Svc" 		then claim_reason = "7"
EMWriteScreen claim_reason, 6, 47
PF3
PF3	'to move past non-inhibiting warning messages on IMIG
EMReadScreen IMIG_screen, 4, 2, 46		'if inhibiting error exists, this will catch it and instruct the user to update IMIG
msgbox IMIG_screen
If IMIG_screen = "IMIG" then script_end_procedure("An error occurred on the IMIG panel. Please update the panel before using the script again.")



Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("SAVE requested/completed for M", memb_number)
Call write_bullet_and_variable_in_CASE_NOTE("Immigration Status", Immig_status_dropdown)
Call write_bullet_and_variable_in_CASE_NOTE("LPR adjusted from", LPR_status_dropdown)
Call write_bullet_and_variable_in_CASE_NOTE("Date of entry", date_of_entry)
Call write_bullet_and_variable_in_CASE_NOTE("Nationality", nationality_dropdown)
Call write_bullet_and_variable_in_CASE_NOTE("Status verfication", status_verification)
Call write_bullet_and_variable_in_CASE_NOTE("Immigration document received", immig_doc_type)
Call write_variable_in_CASE_NOTE("")
	If not_sponsored = 1 then Call write_variable_in_CASE_NOTE("* No sponsor indicated on SAVE.")
	If sponsored = 1 then
		Call write_variable_in_CASE_NOTE("* Client is sponsored. Sponsor is indicated as " & sponsor_name & sponsor_addr & ".")
		IF sponsor_name_two <> "" THEN Call write_variable_in_CASE_NOTE("* Client is sponsored. Sponsor is indicated as " & sponsor_name_two & sponsor_addr_two & ".")
		IF sponsor_name_three <> "" THEN Call write_variable_in_CASE_NOTE("* Client is sponsored. Sponsor is indicated as " & sponsor_name_three & sponsor_addr_three & ".")
	END IF
If additional_save_check = CHECKED then Call write_variable_in_CASE_NOTE("* Additonal SAVE requested.")
If SAVE_docs_check = CHECKED then Call write_variable_in_CASE_NOTE("*attach a copy of the immigration document to request for SAVE")
Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
'TODO add a email reminder
script_end_procedure("")
