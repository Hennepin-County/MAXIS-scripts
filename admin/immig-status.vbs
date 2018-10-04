'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTION - IMMIGRATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 240           'manual run time in seconds
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
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01"
'actual_date = date & ""

'-----------------------------------------------------------------------------------------------------------------------DIALOG
BeginDialog IMIG_dialog, 0, 0, 366, 300, "Immigration Status"
 EditBox 60, 5, 40, 15, MAXIS_case_number
 EditBox 140, 5, 20, 15, memb_number
 EditBox 210, 5, 40, 15, actual_date
 ButtonGroup ButtonPressed
   PushButton 270, 5, 85, 15, "Non-Citizen Guide ", Noncitzn_button
 DropListBox 60, 35, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"US Citizen", immig_status_dropdown
 DropListBox 60, 55, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"N/A", LPR_status_dropdown
 DropListBox 255, 35, 95, 15, "Select One:"+chr(9)+"SAVE Primary"+chr(9)+"SAVE Secondary"+chr(9)+"Alien Card"+chr(9)+"Passport/Visa"+chr(9)+"Re-Entry Prmt"+chr(9)+"INS Correspondence"+chr(9)+"Other Document"+chr(9)+"Certificate of Naturalization"+chr(9)+"No Ver Prvd", status_verification
 DropListBox 255, 55, 95, 15, "Select One:"+chr(9)+"AA Amerasian"+chr(9)+"EH Ethnic Chinese"+chr(9)+"EL Ethnic Lao"+chr(9)+"HG Hmong"+chr(9)+"KD Kurd"+chr(9)+"SJ Soviet Jew"+chr(9)+"TT Tinh"+chr(9)+"AF Afghanistan"+chr(9)+"BK Bosnia"+chr(9)+"CB Cambodia"+chr(9)+"CH China"+chr(9)+"CU Cuba"+chr(9)+"ES El Salvador"+chr(9)+"ER Eritrea"+chr(9)+"ET Ethiopia"+chr(9)+"GT Guatemala"+chr(9)+"HA Haiti"+chr(9)+"HO Honduras"+chr(9)+"IR Iran"+chr(9)+"IZ Iraq"+chr(9)+"LI Liberia"+chr(9)+"MC Micronesia"+chr(9)+"MI Marshall Islands"+chr(9)+"MX Mexico"+chr(9)+"WA Namibia"+chr(9)+"PK Pakistan"+chr(9)+"RP Philippines"+chr(9)+"PL Poland"+chr(9)+"RO Romania"+chr(9)+"RS Russia"+chr(9)+"SO Somalia"+chr(9)+"SF South Africa"+chr(9)+"TH Thailand"+chr(9)+"VM Vietnam"+chr(9)+"OT All Others", nationality_dropdown
 DropListBox 255, 75, 95, 15, "Select One:"+chr(9)+"Certificate of Naturalization"+chr(9)+"Employment Auth Card (I-776 work permit)"+chr(9)+"I-94 Travel Document"+chr(9)+"I-220 B Order of Supervision"+chr(9)+"LPR Card (I-551 green card)"+chr(9)+"SAVE"+chr(9)+"Other"+chr(9)+"No Ver Prvd", immig_doc_type
 EditBox 305, 95, 45, 15, entry_date
 EditBox 305, 115, 45, 15, status_date
 CheckBox 10, 75, 110, 10, "Emailed HP.immigration?", emailHP_CHECKBOX
 CheckBox 10, 90, 90, 10, "Inital SAVE Completed?", save_CHECKBOX
 CheckBox 10, 105, 145, 10, "Additional SAVE Information Requested?", additional_CHECKBOX
 CheckBox 15, 120, 220, 10, "check here if immig document was attached to additional SAVE?", SAVE_docs_check
 OptionGroup RadioGroup1
    RadioButton 15, 155, 25, 10, "No", not_sponsored
    RadioButton 15, 170, 75, 10, "Yes, sponsored by:", sponsored
  EditBox 85, 190, 70, 15, name_sponsor
  EditBox 220, 190, 125, 15, sponsor_addr
  EditBox 85, 210, 70, 15, name_sponsor_two
  EditBox 220, 210, 125, 15, sponsor_addr_two
  EditBox 85, 230, 70, 15, name_sponsor_three
  EditBox 220, 230, 125, 15, sponsor_addr_three
  EditBox 75, 255, 160, 15, other_notes
  EditBox 75, 275, 160, 15, worker_sig
  ButtonGroup ButtonPressed
    CancelButton 310, 275, 45, 15
    OkButton 260, 275, 45, 15
  Text 165, 10, 40, 10, "Actual Date:"
  Text 255, 100, 45, 10, "Date of Entry:"
  Text 10, 40, 50, 10, "Immig Status:"
  Text 10, 60, 50, 10, "LPR Adj From:"
  Text 200, 40, 50, 10, "Status Verified:"
  Text 190, 60, 60, 10, "Nationality/Nation:"
  Text 195, 80, 55, 10, "Immig Doc Type:"
  Text 105, 10, 30, 10, "Memb #:"
  GroupBox 5, 140, 350, 110, "Sponsored on I-864 Affidavit of Support? (LPR COA CODE: C, CF, CR, CX, F, FX, IF, IR)"
  Text 80, 155, 245, 10, "* If date of entry was prior to 12/19/1997 sponsor information is not needed"
  Text 120, 170, 205, 10, "* If sponsor(s) are currenlty unknown please enter unknown and send request additonal SAVE"
  Text 20, 195, 60, 10, "Name of sponsor:"
  Text 165, 195, 55, 10, "Address/Phone:"
  Text 20, 215, 60, 10, "Name of sponsor:"
  Text 165, 215, 55, 10, "Address/Phone:"
  Text 20, 235, 60, 10, "Name of sponsor:"
  Text 165, 235, 55, 10, "Address/Phone:"
  Text 25, 260, 45, 10, "Other Notes:"
  GroupBox 5, 25, 350, 110, "Immigration Information"
  Text 260, 120, 40, 10, "Status Date:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 280, 65, 10, "Worker Signature:"
EndDialog

BeginDialog addimig_dialog, 0, 0, 296, 115, "Additional Information"
  DropListBox 110, 5, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ss_credits
  DropListBox 235, 5, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", verf_sscredits
  DropListBox 110, 20, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", battered_spouse 'mandatory for undoc'
  DropListBox 235, 20, 55, 12, "Select One:"+chr(9)+"YES"+chr(9)+"NO", battered_spouse_verf
  DropListBox 110, 35, 80, 15, "Select One:"+chr(9)+"Veteran"+chr(9)+"Active Duty"+chr(9)+"Spouse of 1 or 2"+chr(9)+"Child of 1 or 2"+chr(9)+"No Military Stat or Other", military_status
  DropListBox 235, 35, 55, 12, "Select One:"+chr(9)+"YES"+chr(9)+"NO", military_status_verf
  DropListBox 110, 50, 135, 15, "Select One:"+chr(9)+"Hmong During Vietnam War"+chr(9)+"Highland Lao During Vietnam"+chr(9)+"Spouse/Widow of 1 Or 2"+chr(9)+"Dep Child of 1 Or 2"+chr(9)+"Native Amer Born Can/Mex"+chr(9)+"N/A", nation_vietnam
  DropListBox 110, 65, 55, 12, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ESL_ctzn 'mandatory for GA'
  DropListBox 235, 65, 55, 12, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ESL_ctzn_verf
  DropListBox 110, 80, 55, 12, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ESL_skills
  ButtonGroup ButtonPressed
    OkButton 195, 95, 45, 15
    CancelButton 245, 95, 45, 15
  Text 20, 10, 90, 10, "40 Social Security Credits:"
  Text 30, 25, 75, 10, "Battered Spouse/Child:"
  Text 55, 40, 50, 10, "Military Status:"
  Text 5, 55, 100, 10, "Hmong, Lao, Native American:"
  Text 30, 70, 80, 10, "St Prog ESL/Ctzn Coop:"
  Text 25, 85, 80, 10, "FSS ESL/Skills Training:"
  Text 200, 10, 35, 10, "Verified?:"
  Text 200, 25, 35, 10, "Verified?:"
  Text 200, 40, 35, 10, "Verified?:"
  Text 200, 70, 35, 10, "Verified?:"
EndDialog

'-----------------------------------------------------------------------------------------------------------THE SCRIPT
Do
	Do
		err_msg = ""
		Do
			dialog IMIG_dialog
			cancel_confirmation
			If ButtonPressed = Noncitzn_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Immigration_and_Non-Citizens.aspx")
		Loop until ButtonPressed = -1
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF memb_number = "" or IsNumeric(memb_number) = False or len(memb_number) > 2 then err_msg = err_msg & vbNewLine & "* Enter a member number."
		IF immig_status_dropdown = "US Citizen" Then
			err_msg = err_msg & vbNewLine & "* This will delete IMIG, update MEMI & MEMB for this member."
			'IF immig_doc_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of immigration document used."
			IF immig_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of current immigration status."
			EXIT DO
		ELSE
			If isdate(actual_date) = FALSE then err_msg = err_msg & vbnewline & "* You must enter an actual date in the footer month that you are working in and not in the future."
			''"21 Refugee" "22 Asylee""23 Deport/Remove Withheld" "24 LPR" "25 Paroled For 1 Year Or More" "26 Conditional Entry < 4/80" "27 Non-immigrant" "28 Undocumented""50 Other Lawfully Residing""US Citizen", immig_status_dropdown
			If isdate(entry_date) = FALSE then err_msg = err_msg & vbnewline & "* Entry Date is required for all persons." 'with exception of Asylee, Deportation/Removal Withheld, or Undocumented statuses."
		END IF
		IF immig_status_dropdown = "22 Asylee" or immig_status_dropdown = "23 Deport/Remove Withheld" Then
			If isdate(status_date) = FALSE then err_msg = err_msg & vbnewline & "* Status Date is required for persons with Asylee or Deportation/Removal Withheld statuses."
		END IF
			IF immig_status_dropdown <> "28 Undocumented" and save_CHECKBOX= UNCHECKED and additional_CHECKBOX = UNCHECKED then err_msg = err_msg & vbNewLine & "* Please select if a SAVE has been run as it is mandatory."
			'IF immig_status_dropdown = "22 Asylee" or immig_status_dropdown = "23 Deport/Remove Withheld" and isdate(status_date) = FALSE then err_msg = err_msg & vbnewline & "* Status Date is required for persons with Asylee or Deportation/Removal Withheld statuses."
			IF immig_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of current immigration status."
			IF immig_status_dropdown = "24 LPR" and LPR_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of LPR adjusted status."
			IF immig_status_dropdown = "24 LPR" and LPR_status_dropdown = "N/A" then err_msg = err_msg & vbNewLine & "* Please advise of LPR adjusted status."
			IF immig_status_dropdown <> "24 LPR" and LPR_status_dropdown <> "Select One:" and LPR_status_dropdown <> "N/A" then err_msg = err_msg & vbNewLine & "* IMMIGRATION STATUS DOES NOT INDICATE LPR, BUT ADJUSTED STATUS IS INDICATED"
			IF immig_doc_type = "Select One:" and immig_status_dropdown <> "28 Undocumented" then err_msg = err_msg & vbNewLine & "* Please advise of immigration document used."
			'Battered Spouse/Child (Y/N): This field is mandatory for undocumented persons, non-immigrants and other lawfully residing persons.
			IF nationality_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of Nationality or Nation."
			IF sponsored = 1 and name_sponsor = "" then err_msg = err_msg & vbNewLine & "* You indicated a sponsor for this case please complete sponsor information."
			IF status_verification = "Certificate of Naturalization" and immig_status_dropdown <> "US Citizen" THEN err_msg = err_msg & vbNewLine & "* You indicated that you have received Certificate of Naturalization immigration status should be US Citizen."
		'END IF
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
Call MAXIS_footer_month_confirmation			'function that confirms that the current footer month/year is the same as what was selected by the user. If not, it will navigate to correct footer month/year

IF immig_status_dropdown <> "US Citizen" Then
	Do
		Do
			err_msg = ""
			dialog addimig_dialog
			cancel_confirmation
			'IF immig_status_dropdown = "28 Undocumented" or immig_status_dropdown = "27 Non-immigrant" or immig_status_dropdown = "50 Other Lawfully Residing" and battered_spouse = "Select One:" THEN	err_msg = err_msg & vbNewLine & "* Please advise if battered spouse or child is applicable."
			'IF nationality_dropdown = "EL Ethnic Lao" or nationality_dropdown = "HG Hmong" and nation_vietnam = "Select One:" Then err_msg = err_msg & vbNewLine & "* Please advise if client has a status during Vietnam War or is Native American born in Mexico or Canada."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF
'write for 10 '
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
the_month = datepart("m", actual_date)
MAXIS_footer_month = right("00" & the_month, 2)
the_year = datepart("yyyy", actual_date)
MAXIS_footer_year = right("00" & the_year, 2)

CALL convert_date_into_MAXIS_footer_month(actual_date, footer_month, footer_year)

Call navigate_to_MAXIS_screen("STAT", "IMIG")
'Making sure we have the correct IMIG
EMReadScreen panel_number, 1, 2, 78
If panel_number = "0" then script_end_procedure("An IMIG panel does not exist. Please create the panel before running the script again. ")
'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
DO
	EMReadScreen current_panel_number, 1, 2, 73
	IMIG_check = MsgBox("Is this the right IMIG?", vbYesNo +vbQuestion, "Confirmation")
	If IMIG_check = vbYes THEN EXIT DO
	If IMIG_check = vbNo THEN TRANSMIT
	If (IMIG_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure("Unable to find another IMIG. Please review the case, and run the script again if applicable.")
Loop until current_panel_number = panel_number

'Updating the IMIG panel
PF9
EMReadScreen error_check, 2, 24, 2	'making sure we can actually update this case.
error_check = trim(error_check)
If error_check <> "" then script_end_procedure("Unable to update this case. Please review case, and run the script again if applicable.")
'if the client is now a citizen this will delete IMIG and update MEMB and MEMI'
IF immig_status_dropdown = "US Citizen" THEN
	IF status_verification = "SAVE Primary" THEN citizen_status = "IM"
	IF status_verification = "SAVE Secondary" THEN citizen_status = "IM"
	IF status_verification = "Alien Card" THEN citizen_status = "IM"
	IF status_verification = "Passport/Visa" THEN citizen_status = "PV"
	IF status_verification = "Re-Entry Prmt" THEN citizen_status = "PV"
	IF status_verification = "INS Correspondence" THEN citizen_status = "IM"
	IF status_verification = "Other Document" THEN citizen_status = "OT"
	IF status_verification = "Certificate of Naturalization" THEN citizen_status = "NP"
	IF status_verification = "No Ver Prvd" THEN citizen_status = "NO"
	'deleting the panel'
	EMwritescreen "DEL", 20, 71
	TRANSMIT
	'Navigates to MEMI tp update citizenship status'
	Call navigate_to_MAXIS_screen("STAT", "MEMI")
	Emwritescreen memb_number, 20, 76
	TRANSMIT
	PF9
	Call create_MAXIS_friendly_date_with_YYYY(actual_date, 0, 6, 35)
	Emwritescreen "Y", 10, 49
	Emwritescreen citizen_status, 10, 78
	TRANSMIT
	'Navigates to MEMB tp update alien ID'
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	Emwritescreen memb_number, 20, 76
	TRANSMIT
	PF9
	Call clear_line_of_text(15, 68)
	TRANSMIT
ELSE
	Call create_MAXIS_friendly_date_with_YYYY(actual_date, 0, 5, 45)
	EMWriteScreen "N", 7, 47			'Sup evidence Y/N field (defaulted to N during this process)
	'converting the immigration stauts from droplist to the applicable MAXIS coding
	'check to see if we need to clear the page before we write to it'
	immig_status = ""
		If immig_status_dropdown = "21 Refugee" then immig_status = "21"
		If immig_status_dropdown = "22 Asylee" then immig_status = "22"
		If immig_status_dropdown = "23 Deport/Remove Withheld" then immig_status = "23"
		If immig_status_dropdown = "24 LPR" then immig_status = "24"
		If immig_status_dropdown = "25 Paroled For 1 Year Or More" then immig_status = "25"
		If immig_status_dropdown = "26 Conditional Entry < 4/80" then immig_status = "26"
		If immig_status_dropdown = "27 Non-immigrant" then immig_status = "27"
		If immig_status_dropdown = "28 Undocumented" then immig_status = "28"
		If immig_status_dropdown = "50 Other Lawfully Residing" then immig_status = "50"
		EMWriteScreen immig_status, 6, 45

		IF entry_date <> "" THEN Call create_MAXIS_friendly_date_with_YYYY(entry_date, 0, 7, 45)
		IF entry_date = "" THEN Call clear_line_of_text(7, 45)
		IF entry_date = "" THEN Call clear_line_of_text(7, 48)
		IF entry_date = "" THEN Call clear_line_of_text(7, 51)
		IF status_date <> "" THEN Call create_MAXIS_friendly_date_with_YYYY(status_date, 0, 7, 71)
		IF status_date = "" THEN Call clear_line_of_text(7, 71)
		IF status_date = "" THEN Call clear_line_of_text(7, 74)
		IF status_date = "" THEN Call clear_line_of_text(7, 77)
		reminder_date = dateadd("d", 10, date)' FOR APPT DATE'

		LPR_status = ""
		If LPR_status_dropdown = "21 Refugee" then LPR_status = "21"
		If LPR_status_dropdown = "22 Asylee" then LPR_status = "22"
		If LPR_status_dropdown = "23 Deport/Remove Withheld" then LPR_status = "23"
		If LPR_status_dropdown = "24 LPR" then LPR_status = "24"
		If LPR_status_dropdown = "25 Paroled For 1 Year Or More" then LPR_status = "25"
		If LPR_status_dropdown = "26 Conditional Entry < 4/80" then LPR_status = "26"
		If LPR_status_dropdown = "27 Non-immigrant" then LPR_status = "27"
		If LPR_status_dropdown = "28 Undocumented" then LPR_status = "28"
		If LPR_status_dropdown = "50 Other Lawfully Residing" then LPR_status = "50"
		EMWriteScreen LPR_status, 9, 45
		If LPR_status = "" THEN Call clear_line_of_text(9, 45)
		'immig_doc_type "Select One:"+chr(9)+"Certificate of Naturalization"+chr(9)+"Employment Auth Card (I-776 work permit)"+chr(9)+"I-94 Travel Document", +chr(9)+"I-220 B Order of Supervision"+chr(9)+"LPR Card (I-551 green card)"+chr(9)+"SAVE"+chr(9)+"Other"
		nationality_status = ""
		IF nationality_dropdown = "AA Amerasian" THEN nationality_status = "AA"
		IF nationality_dropdown = "EH Ethnic Chinese" THEN nationality_status = "EH"
		IF nationality_dropdown = "EL Ethnic Lao" THEN nationality_status = "EL"
		IF nationality_dropdown = "HG Hmong" THEN nationality_status = "HG"
		IF nationality_dropdown = "KD Kurd" THEN nationality_status = "KD"
		IF nationality_dropdown = "SJ Soviet Jew" THEN nationality_status = "SJ"
		IF nationality_dropdown = "TT Tinh" THEN nationality_status = "TT"
		IF nationality_dropdown = "AF Afghanistan" THEN nationality_status = "AF"
		IF nationality_dropdown = "BK Bosnia" THEN nationality_status = "BK"
		IF nationality_dropdown = "CB Cambodia" THEN nationality_status = "CB"
		IF nationality_dropdown = "CH China Mainland" THEN nationality_status = "CH"
		IF nationality_dropdown = "CU Cuba" THEN nationality_status = "CU"
		IF nationality_dropdown = "ES El Salvador" THEN nationality_status = "ES"
		IF nationality_dropdown = "ER Eritrea" THEN nationality_status = "ER"
		IF nationality_dropdown = "ET Ethiopia" THEN nationality_status = "ET"
		IF nationality_dropdown = "GT Guatemala" THEN nationality_status = "GT"
		IF nationality_dropdown = "HA Haiti" THEN nationality_status = "HA"
		IF nationality_dropdown = "HO Honduras" THEN nationality_status = "HO"
		IF nationality_dropdown = "IR Iran" THEN nationality_status = "IR"
		IF nationality_dropdown = "IZ Iraq" THEN nationality_status = "IZ"
		IF nationality_dropdown = "LI Liberia" THEN nationality_status = "LI"
		IF nationality_dropdown = "MC Micronesia" THEN nationality_status = "MC"
		IF nationality_dropdown = "MI Marshall Islands" THEN nationality_status = "MI"
		IF nationality_dropdown = "MX Mexico" THEN nationality_status = "MX"
		IF nationality_dropdown = "WA Namibia" THEN nationality_status = "WA"
		IF nationality_dropdown = "PK Pakistan" THEN nationality_status = "PK"
		IF nationality_dropdown = "RP Philippines" THEN nationality_status = "RP"
		IF nationality_dropdown = "PL Poland" THEN nationality_status = "PL"
		IF nationality_dropdown = "RO Romania" THEN nationality_status = "RO"
		IF nationality_dropdown = "RS Russia" THEN nationality_status = "RS"
		IF nationality_dropdown = "SO Somalia" THEN nationality_status = "SO"
		IF nationality_dropdown = "SF South Africa" THEN nationality_status = "SF"
		IF nationality_dropdown = "TH Thailand" THEN nationality_status = "TH"
		IF nationality_dropdown = "VM Vietnam" THEN nationality_status = "VM"
		IF nationality_dropdown = "OT All Others" THEN nationality_status = "OT"
		EMWriteScreen nationality_status, 10, 45

		'IF save_CHECKBOX = CHECKED or additional_CHECKBOX = CHECKED THEN
		IF status_verification = "SAVE Primary" THEN verif_status = "S1"
		IF status_verification = "SAVE Secondary" THEN verif_status = "S2"
		IF status_verification = "Alien Card" THEN verif_status = "AL"
		IF status_verification = "Passport/Visa" THEN verif_status = "PV"
		IF status_verification = "Re-Entry Prmt" THEN verif_status = "RE"
		IF status_verification = "INS Correspondence" THEN verif_status = "IM"
		IF status_verification = "Other Document" THEN verif_status = "OT"
		IF status_verification = "No Ver Prvd" THEN verif_status = "NO"
		EMWriteScreen verif_status, 8, 45

		'EMReadScreen id_number, 9, 10, 72
		'IF alien_id_number <> id_number THEN MsgBox "The number enter for ID does not match the number entered in the case note"
		'VERIFICATION OF 40 SOCIAL SECURITY CREDITS IS NOT NEEDED  '
		IF ss_credits <> "Select One:" THEN EmWriteScreen ss_credits, 13, 56
		IF ss_credits = "Select One:" THEN Call clear_line_of_text(13, 56)
		IF verf_sscredits <> "Select One:" THEN EmWriteScreen verf_sscredits, 13, 71
		IF verf_sscredits = "Select One:" THEN Call clear_line_of_text(13, 71)
		IF battered_spouse <> "Select One:" THEN EmWriteScreen battered_spouse, 14, 56  'mandatory for undoc'
		IF battered_spouse = "Select One:" THEN Call clear_line_of_text(14, 56)
	    IF battered_spouse_verf <> "Select One:" THEN EmWriteScreen battered_spouse_verf, 14, 71
		IF battered_spouse_verf = "Select One:" THEN Call clear_line_of_text(14, 71)
		IF military_status = "Select One:" THEN Call clear_line_of_text(15, 56)
		IF military_status <> "Select One:" THEN
			IF military_status = "Veteran" THEN EmWriteScreen "1", 15, 56
			IF military_status = "Active Duty" THEN EmWriteScreen "2", 15, 56
			IF military_status = "Spouse of 1 or 2" THEN EmWriteScreen "3", 15, 56
			IF military_status = "Child of 1 or 2" THEN EmWriteScreen "4", 15, 56
			IF military_status = "No Military Stat or Other" THEN EmWriteScreen "N", 15, 56
		END IF
		IF military_status_verf = "Select One:" THEN Call clear_line_of_text(15, 71)
		IF military_status_verf <> "Select One:" THEN EmWriteScreen military_status_verf, 15, 71
		IF nation_vietnam = "Select One:" THEN Call clear_line_of_text(13, 56)
		IF nation_vietnam <> "Select One:" THEN
			IF nation_vietnam = "Hmong During Vietnam War" THEN EmWriteScreen "01", 13, 56
			IF nation_vietnam = "Highland Lao During Vietnam" THEN EmWriteScreen "02", 13, 56
			IF nation_vietnam = "Spouse/Widow of 1 Or 2"THEN EmWriteScreen "03", 13, 56
			IF nation_vietnam = "Dep Child of 1 Or 2"THEN EmWriteScreen "04", 13, 56
			IF nation_vietnam = "Native Amer Born Can/Mex" THEN EmWriteScreen "05", 13, 56
			IF nation_vietnam = "N/A" THEN EmWriteScreen "  ", 13, 56
		END IF
		IF ESL_ctzn <> "Select One:" THEN EmWriteScreen ESL_ctzn, 17, 56  'mandatory for GA'
		IF ESL_ctzn = "Select One:" THEN Call clear_line_of_text(17, 56)
		IF ESL_ctzn_verf <> "Select One:" THEN EmWriteScreen ESL_ctzn_verf, 17, 71
		IF ESL_ctzn_verf = "Select One:" THEN Call clear_line_of_text(17, 71)
		IF ESL_skills <> "Select One:" THEN EmWriteScreen ESL_skills, 18, 56
		IF ESL_skills = "Select One:" THEN Call clear_line_of_text(18, 56)
		Transmit
		'PF3	'to move past non-inhibiting warning messages on IMIG
		EMReadScreen IMIG_screen, 4, 2, 49		'if inhibiting error exists, this will catch it and instruct the user to update IMIG
		msgbox IMIG_screen
		'If IMIG_screen = "IMIG" then script_end_procedure("An error occurred on the IMIG panel. Please update the panel before using the script again.")
END IF

start_a_blank_CASE_NOTE
IF additional_CHECKBOX = CHECKED THEN
 		Call write_variable_in_case_note("IMIG-Instituted Additional SAVE for M" & memb_number)
ELSEIF save_CHECKBOX = CHECKED THEN
	Call write_variable_in_case_note("IMIG-Initial SAVE Completed for M" & memb_number)
ELSEIF immig_status_dropdown = "US Citizen" THEN
	Call write_variable_in_case_note("SAVE Completed for M" & memb_number & " US Citizen")
ELSEIF immig_status_dropdown = "28 Undocumented" THEN
	Call write_variable_in_case_note("Updated IMIG for M" & memb_number)
END IF
Call write_bullet_and_variable_in_case_note("Immigration Status", immig_status_dropdown)
IF LPR_status_dropdown <> "Select One:" then Call write_bullet_and_variable_in_case_note("LPR adjusted from", LPR_status_dropdown)
IF status_date <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Status date", status_date)
IF date_of_entry <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Date of entry", date_of_entry)
IF nationality_dropdown <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Nationality", nationality_dropdown)
IF status_verification <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Status verification", status_verification)
IF status_verification <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Immigration document received", immig_doc_type)
IF emailHP_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Emailed HP Immigration")
Call write_variable_in_case_note("")
If immig_status_dropdown <> "US Citizen" and sponsored = 1 then
	Call write_variable_in_case_note("* Client is sponsored. Sponsor is indicated as " & sponsor_name & sponsor_addr & ".")
	IF sponsor_name_two <> "" THEN Call write_variable_in_case_note("* Client is sponsored. Second Sponsor is indicated as " & sponsor_name_two & sponsor_addr_two & ".")
	IF sponsor_name_three <> "" THEN Call write_variable_in_case_note("* Client is sponsored. Third Sponsor is indicated as " & sponsor_name_three & sponsor_addr_three & ".")
END IF
If save_CHECKBOX = CHECKED then Call write_variable_in_case_note("* SAVE requested.")
If additional_CHECKBOX = CHECKED then Call write_variable_in_case_note("* Additional SAVE requested.")
If SAVE_docs_check = CHECKED then Call write_variable_in_case_note("* Attached a copy of the immigration document to request for SAVE")
Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
IF ss_credits <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("40 Social Security Credits", ss_credits)
IF battered_spouse <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Battered Spouse/Child", battered_spouse)
IF military_status <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Military Status", military_status)
IF nation_vietnam <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Hmong, Lao, Native American", nation_vietnam)
IF ESL_ctzn <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("St Prog ESL/Ctzn Coop", ESL_ctzn)
IF ESL_skills <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("ESL/Skills Training", ESL_skills)
If Outlook_remider = True then call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)
PF3
'TODO add a email to HP IMIG and reminder for second check on additional request
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'IF additional_CHECKBOX = CHECKED THEN CALL create_outlook_email("HSPH.HPImmigration@hennepin.us", "", MAXIS_case_name & maxis_case_number & " Expedited case to be assigned, transferred to team. " & worker_number & "  EOM.", "", "", TRUE)
''create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
IF additional_CHECKBOX = CHECKED THEN
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "SAVE request check: " & reminder_text & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
	Outlook_remider = True
End if
script_end_procedure("Success! Please review your case notes and IMIG panels to ensure they were updated correctly.")
