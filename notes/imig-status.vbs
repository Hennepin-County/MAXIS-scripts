'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IMIG - STATUS.vbs"
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
Call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("08/31/2020", "Updated email information gathering and output functionality for additional stability.", "Ilse Ferris, Hennepin County")
call changelog_update("08/25/2020", "Added handling to ensure the member number is entered as a 2 digit number for readability.", "Casey Love, Hennepin County")
call changelog_update("07/29/2020", "Updated coding to email HPImmigration and handling for when a client is reported as Lawfully Residing.", "MiKayla Handley, Hennepin County")
call changelog_update("08/07/2019", "Updated coding to update citizenship status and verification at new location due to MEMI panel changes associated with New Spouse Income Policy.", "Ilse Ferris, Hennepin County")
call changelog_update("01/25/2019", "Added a case note only option when a case is inactive.", "MiKayla Handley")
call changelog_update("03/28/2018", "Initial version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
'MEMB_number = "01"
'actual_date = date & ""
'Determines which programs are currently status_checking in the month of application
CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'Checking for PRIV cases.
'EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
'IF priv_check = "PRIV" THEN
'	priv_case_list = priv_case_list & "|" & MAXIS_case_number
'ELSE						'For all of the cases that aren't privileged...
'Setting some variables for the loop
CASH_STATUS = FALSE 'overall variable'
CCA_STATUS = FALSE
DW_STATUS = FALSE 'Diversionary Work Program'
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
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
EMReadScreen cash1_prog_check, 2, 6, 67
EMReadScreen cash2_prog_check, 2, 7, 67
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67
EMReadScreen cca_prog_check, 2, 14, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND"  THEN FS_STATUS = TRUE
IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN GRH_STATUS = TRUE
IF hc_status_check = "ACTV" or hc_status_check = "PEND"  THEN HC_STATUS = TRUE
IF cca_status_check = "ACTV" or cca_status_check = "PEND"  THEN CCA_STATUS = TRUE
'Logic to determine if MFIP is active
If cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "RC" or cash1_prog_check = "MS" THEN
	If cash1_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash1_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash1_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash1_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash1_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash1_status_check = ""     Then CASH_STATUS = FALSE
END IF
If cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "RC" or cash2_prog_check = "MS" THEN
	If cash2_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash2_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash2_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash2_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash2_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash2_status_check = ""     Then CASH_STATUS = FALSE
END IF

'IF CASH_STATUS = FALSE or FS_STATUS = FALSE THEN

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 366, 300, "Immigration Status"
 EditBox 60, 5, 40, 15, MAXIS_case_number
 EditBox 140, 5, 20, 15, MEMB_number
 EditBox 210, 5, 40, 15, actual_date
 ButtonGroup ButtonPressed
   PushButton 270, 5, 85, 15, "Non-Citizen Guide ", Noncitzn_button
 DropListBox 60, 35, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 LPR"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"US Citizen", immig_status_dropdown
 DropListBox 60, 55, 110, 15, "Select One:"+chr(9)+"21 Refugee"+chr(9)+"22 Asylee"+chr(9)+"23 Deport/Remove Withheld"+chr(9)+"24 None"+chr(9)+"25 Paroled For 1 Year Or More"+chr(9)+"26 Conditional Entry < 4/80"+chr(9)+"27 Non-immigrant"+chr(9)+"28 Undocumented"+chr(9)+"50 Other Lawfully Residing"+chr(9)+"N/A", LPR_status_dropdown
 DropListBox 255, 35, 95, 15, "Select One:"+chr(9)+"SAVE Primary"+chr(9)+"SAVE Secondary"+chr(9)+"Alien Card"+chr(9)+"Passport/Visa"+chr(9)+"Re-Entry Prmt"+chr(9)+"INS Correspondence"+chr(9)+"Other Document"+chr(9)+"Certificate of Naturalization"+chr(9)+"No Ver Prvd", status_verification
 DropListBox 255, 55, 95, 15, "Select One:"+chr(9)+"AA Amerasian"+chr(9)+"EH Ethnic Chinese"+chr(9)+"EL Ethnic Lao"+chr(9)+"HG Hmong"+chr(9)+"KD Kurd"+chr(9)+"SJ Soviet Jew"+chr(9)+"TT Tinh"+chr(9)+"AF Afghanistan"+chr(9)+"BK Bosnia"+chr(9)+"CB Cambodia"+chr(9)+"CH China"+chr(9)+"CU Cuba"+chr(9)+"ES El Salvador"+chr(9)+"ER Eritrea"+chr(9)+"ET Ethiopia"+chr(9)+"GT Guatemala"+chr(9)+"HA Haiti"+chr(9)+"HO Honduras"+chr(9)+"IR Iran"+chr(9)+"IZ Iraq"+chr(9)+"LI Liberia"+chr(9)+"MC Micronesia"+chr(9)+"MI Marshall Islands"+chr(9)+"MX Mexico"+chr(9)+"WA Namibia"+chr(9)+"PK Pakistan"+chr(9)+"RP Philippines"+chr(9)+"PL Poland"+chr(9)+"RO Romania"+chr(9)+"RS Russia"+chr(9)+"SO Somalia"+chr(9)+"SF South Africa"+chr(9)+"TH Thailand"+chr(9)+"VM Vietnam"+chr(9)+"OT All Others", nationality_dropdown
 DropListBox 255, 75, 95, 15, "Select One:"+chr(9)+"Certificate of Naturalization"+chr(9)+"Employment Auth Card (I-776 work permit)"+chr(9)+"I-94 Travel Document"+chr(9)+"I-220 B Order of Supervision"+chr(9)+"LPR Card (I-551 green card)"+chr(9)+"SAVE"+chr(9)+"Other"+chr(9)+"No Ver Prvd", immig_doc_type
 EditBox 305, 95, 45, 15, entry_date
 EditBox 305, 115, 45, 15, status_date
 CheckBox 10, 75, 110, 10, "Email HP.immigration?", HP_EMAIL_CHECKBOX
 CheckBox 10, 90, 90, 10, "SAVE Completed?", save_CHECKBOX
 CheckBox 10, 105, 145, 10, "Additional SAVE Information Requested?", additional_CHECKBOX
 CheckBox 15, 120, 220, 10, "check here if immig document was attached to additional SAVE?", SAVE_docs_check
 GroupBox 5, 140, 350, 110, "Sponsored on I-864 Affidavit of Support?          (LPR COA CODE: C, CF, CR, CX, F, FX, IF, IR)"
 Text 15, 150, 245, 10, "* If date of entry was prior to 12/19/1997 sponsor(s) information is not needed"
 Text 15, 160, 315, 10, "* If sponsor(s) are currenlty unknown please enter unknown and request an additonal SAVE"
 CheckBox 10, 175, 175, 10, "YES - If sponsored please complete the following:", yes_sponsored
 CheckBox 220, 175, 25, 10, "NO", not_sponsored
 Text 10, 195, 60, 10, "Name of sponsor:"
 EditBox 75, 190, 80, 15, sponsor_name
 Text 165, 195, 55, 10, "Address/Phone:"
 EditBox 220, 190, 125, 15, sponsor_addr
 Text 10, 215, 60, 10, "Name of sponsor:"
 EditBox 75, 210, 80, 15, sponsor_name_two
 Text 165, 215, 55, 10, "Address/Phone:"
 EditBox 220, 210, 125, 15, sponsor_addr_two
 Text 10, 235, 60, 10, "Name of sponsor:"
 EditBox 75, 230, 80, 15, sponsor_name_three
 Text 165, 235, 55, 10, "Address/Phone:"
 EditBox 220, 230, 125, 15, sponsor_addr_three
  EditBox 75, 255, 160, 15, other_notes
  EditBox 75, 275, 160, 15, worker_signature
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
  Text 25, 260, 45, 10, "Other Notes:"
  GroupBox 5, 25, 350, 110, "Immigration Information"
  Text 260, 120, 40, 10, "Status Date:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 280, 65, 10, "Worker Signature:"
  CheckBox 245, 255, 110, 10, "Check here if case is INACTIVE", case_note_only_checkbox
EndDialog

'-----------------------------------------------------------------------------------------------------------THE SCRIPT
Do
	Do
		err_msg = ""
		Do
			dialog Dialog1
			cancel_confirmation
			If ButtonPressed = Noncitzn_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/Immigration_and_Non-Citizens.aspx")
		Loop until ButtonPressed = -1
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF MEMB_number = "" or IsNumeric(MEMB_number) = False or len(MEMB_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a member number."
		IF immig_status_dropdown = "US Citizen" Then
			err_msg = err_msg & vbNewLine & "* This will delete IMIG, SPON, and update MEMI & MEMB for this member."
			EXIT DO
		ELSE
			IF immig_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of current immigration status."
			IF immig_doc_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of immigration document used."
			IF (immig_doc_type = "Certificate of Naturalization" and immig_status_dropdown <> "US Citizen") THEN err_msg = err_msg & vbNewLine & "* You indicated that you have received Certificate of Naturalization immigration status should be US Citizen."
			If isdate(actual_date) = FALSE then err_msg = err_msg & vbnewline & "* You must enter an actual date in the footer month that you are working in and not in the future."
			If isdate(entry_date) = FALSE then err_msg = err_msg & vbnewline & "* Entry Date is required for all persons." 'with exception of Asylee, Deportation/Removal Withheld, or Undocumented statuses."
		END IF
		IF immig_status_dropdown = "27 Non-immigrant" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* You selected that the client is a Non-immigrant please specify status in other notes."
		IF immig_status_dropdown = "28 Undocumented" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* You selected that the client is Undocumented please specify status in other notes."
		IF immig_status_dropdown = "50 Other Lawfully Residing" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* You selected that the client is Lawfully Residing please specify status in other notes."
		IF LPR_status_dropdown = "27 Non-immigrant" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* You selected that the client was a Non-immigrant please specify status in other notes."
		IF LPR_status_dropdown = "28 Undocumented" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* You selected that the client was Undocumented please specify status in other notes."
		IF LPR_status_dropdown = "50 Other Lawfully Residing" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* You selected that the client was Lawfully Residing please specify status in other notes."
		IF immig_status_dropdown = "22 Asylee" or immig_status_dropdown = "23 Deport/Remove Withheld" Then
			If isdate(status_date) = FALSE then err_msg = err_msg & vbnewline & "* Status Date is required for persons with Asylee or Deportation/Removal Withheld statuses."
		END IF
		IF LPR_status_dropdown = "23 Deport/Remove Withheld" or LPR_status_dropdown = "22 Asylee" Then
			If isdate(status_date) = FALSE then err_msg = err_msg & vbnewline & "* Enter the date status was granted as an Asylee or the date Deportation/Removal Withheld was granted."
		END IF
		IF nationality_dropdown = "OT All Others" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please advise of Nationality in Other Notes."
		IF immig_status_dropdown <> "28 Undocumented" and save_CHECKBOX= UNCHECKED and additional_CHECKBOX = UNCHECKED then err_msg = err_msg & vbNewLine & "* Please select if a SAVE has been run as it is mandatory."
		IF immig_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of current immigration status."
		IF immig_status_dropdown = "24 LPR" and LPR_status_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of LPR adjusted status."
		IF immig_status_dropdown = "24 LPR" and yes_sponsored = UNCHECKED and not_sponsored = UNCHECKED then err_msg = err_msg & vbNewLine & "* Please advise of LPR Sponsor Information."
		IF immig_status_dropdown = "24 LPR" and LPR_status_dropdown = "N/A" then err_msg = err_msg & vbNewLine & "* Please advise of LPR adjusted status."
		IF immig_status_dropdown <> "24 LPR" and LPR_status_dropdown <> "Select One:" and LPR_status_dropdown <> "N/A" then err_msg = err_msg & vbNewLine & "* Immigration status does not indicate LPR, but adjusted status is indicated."
		IF immig_doc_type = "Select One:" and immig_status_dropdown <> "28 Undocumented" then err_msg = err_msg & vbNewLine & "* Please advise of immigration document used."
		'Battered Spouse/Child (Y/N): This field is mandatory for undocumented persons, non-immigrants and other lawfully residing persons.
		IF nationality_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "* Please advise of Nationality or Nation."
		IF yes_sponsored = CHECKED and sponsor_name = "" then err_msg = err_msg & vbNewLine & "* You indicated a sponsor for this case please complete sponsor information."
		IF yes_sponsored = CHECKED and not_sponsored = CHECKED then err_msg = err_msg & vbNewLine & "* You indicated a sponsor for this case please complete sponsor information and uncheck no."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

IF case_note_only_checkbox <> CHECKED THEN
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 370, 240, "Additional Information for IMIG"
      DropListBox 115, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ss_credits
      DropListBox 300, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ss_credits_verf
      DropListBox 115, 30, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", battered_spouse 'mandatory for undoc'
      DropListBox 300, 30, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", battered_spouse_verf
      DropListBox 115, 45, 80, 15, "Select One:"+chr(9)+"Veteran"+chr(9)+"Active Duty"+chr(9)+"Spouse of 1 or 2"+chr(9)+"Child of 1 or 2"+chr(9)+"No Military Stat or Other", military_status
      DropListBox 300, 45, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", military_status_verf
      DropListBox 115, 60, 135, 15, "Select One:"+chr(9)+"Hmong During Vietnam War"+chr(9)+"Highland Lao During Vietnam"+chr(9)+"Spouse/Widow of 1 Or 2"+chr(9)+"Dep Child of 1 Or 2"+chr(9)+"Native Amer Born Can/Mex"+chr(9)+"N/A", nation_vietnam
      DropListBox 115, 75, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ESL_ctzn 'mandatory for GA'
      DropListBox 300, 75, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ESL_ctzn_verf
      DropListBox 115, 90, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ESL_skills
      ButtonGroup ButtonPressed
    	OkButton 260, 100, 45, 15
    	CancelButton 310, 100, 45, 15
      GroupBox 5, 5, 360, 115, "IMIG fields"
      Text 20, 20, 90, 10, "40 Social Security Credits:"
      Text 260, 20, 35, 10, "Verified?:"
      Text 30, 35, 75, 10, "Battered Spouse/Child:"
      Text 260, 35, 35, 10, "Verified?:"
      Text 55, 50, 50, 10, "Military Status:"
      Text 10, 170, 345, 15, "* Military Status: If the client is an active duty member or veteran of the US armed forces (or a spouse or unmarried minor child) of an active duty member or veteran."
      Text 30, 80, 80, 10, "St Prog ESL/Ctzn Coop:"
      Text 260, 50, 35, 10, "Verified?:"
      Text 260, 80, 35, 10, "Verified?:"
      GroupBox 5, 120, 360, 120, "Information: Please see instructions for additional information"
      Text 10, 145, 350, 15, "* Battered Spouse/Child: Mandatory field only for those LPRs not otherwise eligible for federal funding. HSR records whether the person has 40 Social Security work credits."
      Text 30, 95, 80, 10, "FSS ESL/Skills Training:"
      Text 10, 195, 345, 20, "* Hmong, Lao, Native American: If the non-citizen client is Hmong, Lao, or Native American born in Canada, please PF12 on the line for more information."
      Text 10, 220, 345, 15, "* St Prog ESL/Ctzn Coop: This field needs to be completed for all LPRs age 18 through 69 in the GA or state-funded MFIP unit."
      Text 10, 65, 100, 10, "Hmong, Lao, Native American:"
      Text 10, 130, 350, 10, "* 40 Social Security Credits: Mandatory field only for those LPRs not otherwise eligible for federal funding. "
    EndDialog
    IF immig_status_dropdown <> "US Citizen" Then
    	Do
    		Do
    			err_msg = ""
    			dialog Dialog1
    			cancel_confirmation
    			IF battered_spouse = "Select One:" THEN
    			 	IF immig_status_dropdown = "28 Undocumented" or immig_status_dropdown = "27 Non-immigrant" or immig_status_dropdown = "50 Other Lawfully Residing" THEN err_msg = err_msg & vbNewLine & "* Please advise if battered spouse or child is applicable."
    			END IF
    			IF nation_vietnam = "Select One:" THEN
    				IF nationality_dropdown = "EL Ethnic Lao" or nationality_dropdown = "HG Hmong"  or nationality_dropdown = "HG Hmong" THEN err_msg = err_msg & vbNewLine & "* Please advise if client has a status during Vietnam War or is Native American born in Mexico or Canada."
    			END IF
    			IF yes_sponsored = CHECKED and CASH_STATUS = TRUE THEN
    				IF ss_credits <> "Select One:" and ss_credits_verf = "Select One:" THEN err_msg = err_msg & vbNewLine & "* You selected that social secuirty credits are verified, please advise if SS credits are applicable."
    			END IF
    			If CASH_STATUS = TRUE THEN
    				IF ESL_ctzn = "Select One:" and immig_status_dropdown = "24 LPR" THEN err_msg = err_msg & vbNewLine & "* Please advise of ESL/Citizenship requirements for state funded cash, see CM 11.03.03 and CM 11.03.09"
    			END IF
    			IF battered_spouse <>  "Select One:" and battered_spouse_verf = "Select One:" THEN err_msg = err_msg & vbNewLine & "* You selected that battered spouse is verified, please advise if advise if battered spouse is applicable."
    			IF military_status <> "Select One:" and military_status_verf  = "Select One:" THEN err_msg = err_msg & vbNewLine & "* You selected that military status is verified, please advise if military status is applicable."
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    		LOOP UNTIL err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
    END IF

    'Defaults the date status_checked to today
    status_checked_date = date & ""
    Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
    the_month = datepart("m", actual_date)
    MAXIS_footer_month = right("00" & the_month, 2)
    the_year = datepart("yyyy", actual_date)
    MAXIS_footer_year = right("00" & the_year, 2)
    CALL convert_date_into_MAXIS_footer_month(actual_date, footer_month, footer_year)

    Call navigate_to_MAXIS_screen("STAT", "IMIG")
    EmWriteScreen MEMB_number, 20, 76
    TRANSMIT
    'Making sure we have the correct IMIG
    EMReadScreen panel_number, 1, 2, 78
    If panel_number = "0" then script_end_procedure("An IMIG panel does not exist. Please create the panel before running the script again. ")
    'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
    EMReadScreen current_panel_check, 4, 2, 49
    IF current_panel_check = "IMIG" THEN
    	DO
    		EMReadScreen current_panel_number, 1, 2, 73
    		IMIG_check = MsgBox("Is this the right IMIG?", vbYesNo +vbQuestion, "Confirmation")
    		If IMIG_check = vbYes THEN EXIT DO
    		If IMIG_check = vbNo THEN
    			EmWriteScreen MEMB_number, 20, 76
    			TRANSMIT
    		END IF
    		If (IMIG_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure("Unable to find another IMIG. Please review the case, and run the script again if applicable.")
    	Loop until current_panel_number = panel_number
    ELSE
    	back_to_self
    	Call navigate_to_MAXIS_screen("STAT", "IMIG")
    	EmWriteScreen MEMB_number, 20, 76
    	TRANSMIT
    END IF
    '-------------------------------------------------------------------------------Updating the IMIG panel
    PF9
    EMReadScreen alien_id_number, 9, 10, 72
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
    	'Navigates to delete SPON'
    	Call navigate_to_MAXIS_screen("STAT", "SPON")
    	Emwritescreen MEMB_number, 20, 76
    	TRANSMIT
    	'Making sure we have the correct IMIG
    	EMReadScreen panel_number, 1, 2, 78
    	If panel_number = "0" then
    		EMReadScreen error_msg, 2, 24, 2
    		error_msg = TRIM(error_msg)
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	else
    		PF9
    		EMwritescreen "DEL", 20, 71
    		TRANSMIT
    	end if
    	'Navigates to MEMI tp update citizenship status'
    	Call navigate_to_MAXIS_screen("STAT", "MEMI")
    	Emwritescreen MEMB_number, 20, 76
    	TRANSMIT
    	PF9
    	Call create_MAXIS_friendly_date_with_YYYY(actual_date, 0, 6, 35)
    	Emwritescreen "Y", 11, 49
    	Emwritescreen citizen_status, 11, 78
    	TRANSMIT
    	'Navigates to MEMB tp update alien ID'
    	Call navigate_to_MAXIS_screen("STAT", "MEMB")
    	Emwritescreen MEMB_number, 20, 76
    	TRANSMIT

    	EMReadScreen alien_id_number, 9, 15, 68
    	IF alien_id_number <> "" THEN
    		PF9
    		Call clear_line_of_text(15, 68)
    		TRANSMIT
    	END IF
    ELSE
    	Call create_MAXIS_friendly_date_with_YYYY(actual_date, 0, 5, 45)
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
		Call change_date_to_soonest_working_day(reminder_date, "BACK")

    	LPR_status = ""
    	If LPR_status_dropdown = "21 Refugee" then LPR_status = "21"
    	If LPR_status_dropdown = "22 Asylee" then LPR_status = "22"
    	If LPR_status_dropdown = "23 Deport/Remove Withheld" then LPR_status = "23"
    	If LPR_status_dropdown = "24 None" then LPR_status = "24"
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

    	IF status_verification = "SAVE Primary" THEN verif_status = "S1"
    	IF status_verification = "SAVE Secondary" THEN verif_status = "S2"
    	IF status_verification = "Alien Card" THEN verif_status = "AL"
    	IF status_verification = "Passport/Visa" THEN verif_status = "PV"
    	IF status_verification = "Re-Entry Prmt" THEN verif_status = "RE"
    	IF status_verification = "INS Correspondence" THEN verif_status = "IM"
    	IF status_verification = "Other Document" THEN verif_status = "OT"
    	IF status_verification = "No Ver Prvd" THEN verif_status = "NO"
    	EMWriteScreen verif_status, 8, 45

    	'EMReadScreen alien_id_number, 9, 10, 72
    	'IF alien_id_number <> id_number THEN MsgBox "The number enter for ID does not match the number entered in the case note"
    	'VERIFICATION OF 40 SOCIAL SECURITY CREDITS IS NOT NEEDED  '
    	IF ss_credits = "Select One:" THEN Call clear_line_of_text(13, 56)
    	IF ss_credits_verf = "Select One:" THEN Call clear_line_of_text(13, 71)

    	IF battered_spouse = "Select One:" THEN Call clear_line_of_text(14, 56)
    	IF battered_spouse_verf = "Select One:" THEN Call clear_line_of_text(14, 71)

    	IF military_status = "Select One:" THEN Call clear_line_of_text(15, 56)
    	IF military_status_verf = "Select One:" THEN Call clear_line_of_text(15, 71)

    	IF nation_vietnam = "Select One:" THEN Call clear_line_of_text(16, 56)

    	IF ESL_ctzn = "Select One:" THEN Call clear_line_of_text(17, 56)
    	IF ESL_ctzn_verf = "Select One:" THEN Call clear_line_of_text(17, 71)
    	IF ESL_skills = "Select One:" THEN Call clear_line_of_text(18, 56)

    	IF ss_credits <> "Select One:" THEN EmWriteScreen ss_credits, 13, 56
    	IF ss_credits_verf <> "Select One:" THEN EmWriteScreen ss_credits_verf, 13, 71

    	IF battered_spouse <> "Select One:" THEN EmWriteScreen battered_spouse, 14, 56  'mandatory for undoc'
    	IF battered_spouse_verf <> "Select One:" THEN EmWriteScreen battered_spouse_verf, 14, 71

    	IF military_status <> "Select One:" THEN
    		IF military_status = "Veteran" THEN EmWriteScreen "1", 15, 56
    		IF military_status = "Active Duty" THEN EmWriteScreen "2", 15, 56
    		IF military_status = "Spouse of 1 or 2" THEN EmWriteScreen "3", 15, 56
    		IF military_status = "Child of 1 or 2" THEN EmWriteScreen "4", 15, 56
    		IF military_status = "No Military Stat or Other" THEN EmWriteScreen "N", 15, 56
    	END IF
    	IF military_status_verf <> "Select One:" THEN EmWriteScreen military_status_verf, 15, 71

    	IF nation_vietnam <> "Select One:" THEN
    		IF nation_vietnam = "Hmong During Vietnam War" THEN EmWriteScreen "01", 16, 56
    		IF nation_vietnam = "Highland Lao During Vietnam" THEN EmWriteScreen "02", 16, 56
    		IF nation_vietnam = "Spouse/Widow of 1 Or 2"THEN EmWriteScreen "03", 16, 56
    		IF nation_vietnam = "Dep Child of 1 Or 2"THEN EmWriteScreen "04", 16, 56
    		IF nation_vietnam = "Native Amer Born Can/Mex" THEN EmWriteScreen "05", 16, 56
    		IF nation_vietnam = "N/A" THEN Call clear_line_of_text(16, 56)
    	END IF

    	IF ESL_ctzn <> "Select One:" THEN EmWriteScreen ESL_ctzn, 17, 56  'mandatory for GA'
    	IF ESL_ctzn_verf <> "Select One:" THEN EmWriteScreen ESL_ctzn_verf, 17, 71
    	IF ESL_skills <> "Select One:" THEN EmWriteScreen ESL_skills, 18, 56
    	Transmit
    END IF
END IF
start_a_blank_CASE_NOTE
IF additional_CHECKBOX = CHECKED THEN
	Call write_variable_in_case_note("IMIG-Instituted Additional SAVE for M" & MEMB_number)
ELSEIF save_CHECKBOX = CHECKED THEN
	Call write_variable_in_case_note("IMIG-SAVE Completed for M" & MEMB_number)
ELSEIF immig_status_dropdown = "US Citizen" THEN
	Call write_variable_in_case_note("IMIG-SAVE Completed for M" & MEMB_number & " US Citizen")
	Call write_variable_in_case_note("* Updated MEMB to remove Alien ID")
	Call write_variable_in_case_note("* Updated MEMI to correct status")
	Call write_variable_in_case_note("* Deleted IMIG and SPON")
	Call write_variable_in_case_note("* Sent status verification to resident.")
ELSEIF immig_status_dropdown = "28 Undocumented" THEN
	Call write_variable_in_case_note("Updated IMIG for M" & MEMB_number)
END IF
Call write_bullet_and_variable_in_case_note("Immigration Status", immig_status_dropdown)
IF LPR_status_dropdown <> "Select One:" then Call write_bullet_and_variable_in_case_note("LPR adjusted from", LPR_status_dropdown)
IF status_date <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Status date", status_date)
IF date_of_entry <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Date of entry", date_of_entry)
IF nationality_dropdown <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Nationality", nationality_dropdown)
IF status_verification <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Status verification", status_verification)
IF status_verification <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Immigration document received", immig_doc_type)
IF HP_EMAIL_CHECKBOX = CHECKED THEN Call write_variable_in_case_note("* Emailed HP Immigration")
Call write_variable_in_case_note("")
If yes_sponsored = CHECKED then
	Call write_variable_in_case_note("* Client is sponsored")
	Call write_bullet_and_variable_in_case_note("Sponsor is indicated as", sponsor_name)
	Call write_bullet_and_variable_in_case_note("Sponsor address is", sponsor_addr)
	IF sponsor_name_two <> "" THEN
		Call write_variable_in_case_note("---")
		Call write_bullet_and_variable_in_case_note("Second sponsor is indicated as", sponsor_name_two)
		Call write_bullet_and_variable_in_case_note("Second sponsor address is", sponsor_addr_two)
	END IF
	IF sponsor_name_three <> "" THEN
		Call write_variable_in_case_note("---")
		Call write_bullet_and_variable_in_case_note("Third sponsor is indicated as", sponsor_name_three)
		Call write_bullet_and_variable_in_case_note("Third sponsor address is", sponsor_addr_three)
	END IF
ELSE Call write_variable_in_case_note("* No Sponsor indicated or sponsor is not applicable")
END IF
If save_CHECKBOX = CHECKED then Call write_variable_in_case_note("* SAVE Completed and sent to case file.")
If additional_CHECKBOX = CHECKED then
	Call write_variable_in_case_note("* Additional SAVE requested")
 	Call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
END IF
If SAVE_docs_check = CHECKED then Call write_variable_in_case_note("* Attached a copy of the immigration document to request for SAVE")
Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
IF ss_credits <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("40 Social Security Credits", ss_credits)
IF battered_spouse <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Battered Spouse/Child", battered_spouse)
IF military_status <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("Military Status", military_status)
IF nationality_dropdown = "EL Ethnic Lao" or nationality_dropdown = "HG Hmong" THEN Call write_bullet_and_variable_in_case_note("Status during Vietnam War", nation_vietnam)
IF nation_vietnam = "Native Amer Born Can/Mex" THEN Call write_bullet_and_variable_in_case_note("Native American born in Mexico or Canada", nation_vietnam)
IF ESL_ctzn <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("St Prog ESL/Ctzn Coop", ESL_ctzn)
IF ESL_skills <> "Select One:" THEN Call write_bullet_and_variable_in_case_note("ESL/Skills Training", ESL_skills)
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)
PF3
IF HP_EMAIL_CHECKBOX = CHECKED THEN
    message_array = "CASE NOTE" & vbcr
	EMWriteScreen "x", 5, 3
	TRANSMIT
	note_row = 4			'Beginning of the case notes
	Do 						'Read each line
		EMReadScreen note_line, 76, note_row, 3
		note_line = trim(note_line)
		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
		message_array = message_array & note_line & vbcr		'putting the lines together
		note_row = note_row + 1
		If note_row = 18 then 									'End of a single page of the case note
			EMReadScreen next_page, 7, note_row, 3
			If next_page = "More: +" Then 						'This indicates there is another page of the case note
				PF8												'goes to the next line and resets the row to read'\
				note_row = 4
			End If
		End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages

    email_header = "Please review for accuracy: Case #" & MAXIS_case_number & ", Member #" & memb_number
    'Function create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
	Call create_outlook_email("", "HP.Immigration@hennepin.us", "", "", email_header, 1, False, "", "", False, "", message_array, False, "", FALSE)
END IF
''create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
IF additional_CHECKBOX = CHECKED THEN
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "SAVE request reminder for " & MAXIS_case_number, "", "", TRUE, 5, "")
	Outlook_remider = True
End if
script_end_procedure_with_error_report("Please review your case notes, email, and IMIG panel to ensure accuracy.")
