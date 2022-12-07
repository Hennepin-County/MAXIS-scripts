'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    HOW THE DAIL SCRUBER WORKS:
'
'    This script opens up other script files, using a custom function (run_DAIL_scrubber_script), followed by the path to the script file. It's done this
'      way because there could be hundreds of DAIL messages, and to work all of the combinations into one script would be incredibly tedious and long.
'
'    This script works by moving the message (where the cursor is located) to the top of the screen, and then reading the message text. Whatever the
'      message text says dictates which script loads up.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - DAIL SCRUBBER.vbs"
start_time = timer

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
Call changelog_update("11/14/2022", "Added DAIL Scrubber support for additional ABAWD/WREG messages.", "Ilse Ferris, Hennepin County")
Call changelog_update("05/23/2022", "Added DAIL Scrubber support for SVES covered quarters DAIL messages.", "Ilse Ferris, Hennepin County")
Call changelog_update("04/17/2020", "DAILs for COLA - Review and Approve can now call Approved Programs or Closed Programs if the approval is not for Health Care.", "Casey Love, Hennepin County")
Call changelog_update("06/13/2019", "Added support for the following COLA message: CLAIM NUMBER XXXXXXXXXX NOT MATCHED - REVIEW CLAIM NUMBER AND CORRECT UNEA", "Ilse Ferris, Hennepin County")
Call changelog_update("06/13/2019", "Added DAIL messages for JULY COLA to run the COLA Review and Approve option. See instructions for full detail of messages now handled.", "Casey Love, Hennepin County")
call changelog_update("5/31/2019", "The DAIL message for COLA Review and Approve now has specific handling to either review or approve Health Care eligibility. (Additional programs to be added at a later date.)", "Casey Love, Hennepin County")
call changelog_update("4/26/2019", "The DAIL messages for Over Due Baby, Incarceration, and additional enhancements to handle for other messages has been added.", "MiKayla Handley, Hennepin County")
call changelog_update("4/9/2019", "The DAIL message for Student Income ending has changed. Updated the script to know the new message.", "Casey Love, Hennepin County")
call changelog_update("10/18/2018", "Updated to support updated ABAWD message 'SNAP ABAWD ELIGIBILITY HAS EXPIR'.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

' Function to PF7 all the way to the top of the DAIL list '
function go_to_top_of_dail()
	Do
		PF7							'scroll up'
		first_page, 19, 24, 10		'read for amessage at the bottom of the panel indicating we are at the top
	Loop until first_page = "ONLY SCROLL FORWARD"		'once the messages is displayed, the function is done.
end function

'funcction to find the message the script was started on and bring it to the top of the DAILs.
function bring_correct_message_to_top()
'The commparison varibales are read at the beginning of the DAIL scrubber and should not be redenined:
	'dail_pers_indicator
	'dail_case_indicator
	'find_msg_details
	Do
		EMReadscreen dail_check, 4, 2, 48			'making sure we are at the DAIL
		If dail_check <> "DAIL" then PF3			''backing out to get to the DAIL

		EMReadScreen self_check, 4, 2, 50			'If wew go to far and are at SELF - we are going to navvigate back to DAIL
		If self_check = "SELF" Then
			Call navigate_to_MAXIS_screen("DAIL", "DAIL")
			Exit Do
		End If
	Loop until dail_check = "DAIL"

	'Here we look for the right header but person informmmation and case information
	header_row = 4
	Do
		header_row = header_row + 1
		EMReadScreen dail_pers_header, len(dail_pers_indicator), header_row, 5	'reading the person header'
		EMReadScreen dail_case_header, 18, header_row, 63						'reading the case header'
		If header_row = 19 Then													'if we are at the end of a page of DAILs, we need to go to the next one.
			PF8
			EMReadScreen last_page, 9, 24, 14
			If last_page = "LAST PAGE" Then call go_to_top_of_dail
			header_row = 5
		End If
	Loop until dail_pers_header = dail_pers_indicator and dail_case_header = dail_case_indicator	'stopping when we get to the right header

	dail_row = header_row + 1													'the dail mmessages start one row below the header
	EMReadScreen read_each_dail, 76, dail_row, 5								'read the whole DAIL line
	Do while read_each_dail <> find_msg_details									'comparing each line and stopping when we have found the one we started on.
		dail_row = dail_row + 1
		If dail_row = 19 Then
			PF8
			dail_row = 6
		End If
		EMReadScreen read_each_dail, 77, dail_row, 4
	Loop
	Call write_value_and_transmit("T", dail_row, 3)								'bringing the correct one to the TOP of the DAIL page
end function

'CONNECTS TO DEFAULT SCREEN
EMConnect ""
match_found = FALSE

'Finding the top of this case's list of dails.
EMGetCursor dail_row, dail_col
EMReadScreen find_msg_details, 76, dail_row, 5		'this is the WHOLE line - with type and footer month - we will use this to get back to it if needed

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'now we are reading the header information to make sure we get to the right person and case if we ever have to navigate away from and back to the DAIL
header_row = dail_row-1
EMReadScreen header_in_row_above, 3, header_row, 59
Do While header_in_row_above <> "-->"
	header_row = header_row -1
	EMReadScreen header_in_row_above, 3, header_row, 59
Loop
EMReadScreen dail_pers_indicator, 55, header_row, 5								'this is the information of the case name AND inidivvidual name if it is indicated in ()'
EMReadScreen dail_case_indicator, 18, header_row, 63							'this is the header with case number, including the CASE NUMBER words
'here we are removing the trailing '-' from the person header information
Do
	If right(dail_pers_indicator, 1) = "-" Then dail_pers_indicator = left(dail_pers_indicator, len(dail_pers_indicator)-1)
Loop until right(dail_pers_indicator, 1) <> "-"

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
Call write_value_and_transmit("T", dail_row, dail_col)
EMReadScreen MAXIS_check, 5, 1, 39												'checking to mmake sure wew are not passworded out.'
If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
	call script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
End If

'The following reads the message in full for the end part (which tells the worker which message was selected)
EMReadScreen full_message, 60, 6, 20
full_message = trim(full_message)
EmReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

' 'This functionalitty will identify if the script is being run in INQUIRY
' 'Cannot tell if Production or Training, but that is less important
' running_in_INQUIRY = False							'default to false'
' EMSendKey "W"										'attempt to navigate to DAIL/WRIT - wwhich requires write access
' TRANSMIT
' EMReadScreen cannot_access_msg, 17, 24, 2			'reading for the message that states you cannot access WRIT
' If cannot_access_msg = "YOU CANNOT ACCESS" Then		'If found we are most likely in INQUIRY (could be PRIV or Out of County)
' 	running_in_INQUIRY = True						'setting the boolean to true
' 	EMSendKey " "									'blanking out the 'WRIT' attempt
' Else
' 	PF3												'if the message is NOT found, we are in WRIT and need to back out
' End If
'
' call bring_correct_message_to_top

'THE FOLLOWING CODES ARE THE INDIVIDUAL MESSAGES. IT READS THE MESSAGE, THEN CALLS A NEW SCRIPT.----------------------------------------------------------------------------------------------------

'Random messages generated from an affiliated case (loads AFFILIATED CASE LOOKUP) OR XFS Closed for Postponed Verifications (loads POSTPONTED XFS VERIFICATIONS)
'Both of these messages start with 'FS' on the DAIL, so they need to be nested, or it never gets passed the affilated case look up
EMReadScreen stat_check, 4, 6, 6
If stat_check = "FS  " or stat_check = "HC  " or stat_check = "GA  " or stat_check = "MSA " or stat_check = "STAT" then
	'now it checks if you are acctually running from a XFS Autoclosed DAIL. These messages don't have an affiliated case attached - so there will be no overlap
	match_found = TRUE
	EMReadScreen xfs_check, 49, 6, 20
	If xfs_check = "CASE AUTO-CLOSED FOR FAILURE TO PROVIDE POSTPONED" then
		call run_from_GitHub(script_repository & "dail/postponed-expedited-snap-verifications.vbs")
	Else
		call run_from_GitHub(script_repository & "dail/affiliated-case-lookup.vbs")
	End If
End If

'Checking for 12 month contact TIKL from CAF and CAR scripts(loads NOTICES - 12 month contact)
EMReadScreen twelve_mo_contact_check, 57, 6, 20
IF twelve_mo_contact_check = "IF SNAP IS OPEN, REVIEW TO SEE IF 12 MONTH CONTACT LETTER" THEN
	match_found = TRUE
	run_from_GitHub(script_repository & "notices/12-month-contact.vbs")
END IF

'Run NOTES - AVS
If Instr(full_message, "AN UPDATED DHS-7823 - AVS AUTH FORM(S) HAS BEEN REQUESTED") OR _
   Instr(full_message, "AVS 10-DAY CHECK IS DUE") OR _
   Instr(full_message, "DHS-7823 - AVS AUTH FORM(S) HAVE BEEN REQUESTED FOR THIS") then
   match_found = True
   run_from_GitHub(script_repository & "notes/avs.vbs")
 End if

'RSDI/BENDEX info received by agency (loads BNDX SCRUBBER)
If instr(full_message, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/bndx-scrubber.vbs")
END IF

'CIT/ID has been verified through the SSA (loads CITIZENSHIP VERIFIED)
EMReadScreen CIT_check, 46, 6, 20
If CIT_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/citizenship-verified.vbs")
END IF

'COLA REVIEW AND APPROVE RESPONSE
If InStr(full_message, "COLA UPDATES IN STAT COMPLETED. REVIEW AND APPROVE") <> 0 Then review_and_approve_from_COLA = TRUE
If InStr(full_message, "REVIEW MEDICARE SAVINGS PROGRAM ELIGIBILITY FOR POSSIBLE") <> 0 Then review_and_approve_from_COLA = TRUE
If InStr(full_message, "REVIEW HEALTH CARE ELIGIBILITY FOR POSSIBLE CHANGES DUE TO") <> 0 Then review_and_approve_from_COLA = TRUE
If InStr(full_message, "PERSON DOES NOT HAVE AN APPROVED HEALTH CARE BUDGET") <> 0 Then review_and_approve_from_COLA = TRUE
If InStr(full_message, "PERSON HAS MAINTENANCE NEEDS ALLOWANCE - REVIEW MEDICAL") <> 0 Then review_and_approve_from_COLA = TRUE
If InStr(full_message, "REVIEW MA-EPD FOR POSSIBLE PREMIUM CHANGES DUE TO") <> 0 Then review_and_approve_from_COLA = TRUE
If InStr(full_message, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS - REVIEW") <> 0 Then review_and_approve_from_COLA = TRUE

If review_and_approve_from_COLA = TRUE Then
    match_found = TRUE
    Call run_from_GitHub(script_repository & "dail/cola-review-and-approve.vbs")
End If

'COLA SVES RESPONSE
If instr(full_message, "REVIEW SVES RESPONSE") or instr(full_message, "REVIEW CLAIM NUMBER") then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/cola-sves-response.vbs")
END IF

'Disability certification ends in 60 days (loads DISA MESSAGE)
EMReadScreen DISA_check, 58, 6, 20
If DISA_check = "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/disa-message.vbs")
END IF

'EMPS - ES Referral missing
EMReadScreen EMPS_ES_check, 52, 6, 20
If EMPS_ES_check = "EMPS:ES REFERRAL DATE IS BLANK FOR NON-EXEMPT PERSON" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/es-referral-missing.vbs")
END IF

'EMPS - Financial Orientation date needed
EMReadScreen EMPS_Fin_Ori_check, 57, 6, 20
If EMPS_Fin_Ori_check = "REVIEW EMPS PANEL FOR FINANCIAL ORIENT DATE OR GOOD CAUSE" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/mfip-orientation-dail.vbs")
END IF

'Client can receive an FMED deduction for SNAP (loads FMED DEDUCTION)
EMReadScreen FMED_check, 59, 6, 20
If FMED_check = "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/fmed-deduction.vbs")
END IF

'Remedial care messages. May only happen at COLA (loads LTC - REMEDIAL CARE)
EMReadScreen remedial_care_check, 41, 6, 20
If remedial_care_check = "REF 01 PERSON HAS REMEDIAL CARE DEDUCTION" then
	match_found = TRUE
	CALL run_from_GitHub(script_repository & "dail/ltc-remedial-care.vbs")
END IF
'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 15, 6, 20
If HIRE_check = "NEW JOB DETAILS" or left(HIRE_check, 4) = "SDNH" then
    match_found = TRUE
	call run_from_GitHub(script_repository & "dail/new-hire.vbs")
END IF
'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 11, 6, 37
EmReadscreen fed_match, 4, 6, 20        'SDNH can use the same string to review, NDNH cannot (of course)
If HIRE_check = "JOB DETAILS" or left(fed_match, 4) = "NDNH" then
	match_found = TRUE
    call run_from_GitHub(script_repository & "dail/new-hire-ndnh.vbs")
END IF
'federal prisoner register support messages
EMReadScreen ISPI_check, 4, 6, 6
If ISPI_check = "ISPI" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/incarceration.vbs")
END IF

'MEMBER HAS BEEN DISABLED 2 YEARS - REFER TO MEDICARE
EMReadScreen MEDI_check, 52, 6, 20
If MEDI_check = "MEMBER HAS BEEN DISABLED 2 YEARS - REFER TO MEDICARE" then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/medi-check.vbs")
END IF

'Sends NOMI is DAIL generated by the REVS scrubber (loads SEND NOMI)
EMReadScreen paperless_check, 8, 6, 20
If paperless_check = "%^% SENT" then
	match_found = TRUE
	run_from_DAIL = TRUE
    call run_from_GitHub(script_repository &  "dail/paperless-dail.vbs")
End If

'SSI info received by agency (loads SDX INFO HAS BEEN STORED)
If instr(full_message, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") then
    match_found = TRUE
	call run_from_GitHub(script_repository & "dail/sdx-info-has-been-stored.vbs")
END IF

'SSA info received by agency (loads TPQY RESPONSE)
If instr(full_message, "TPQY RESPONSE RECEIVED FROM SSA") or instr(full_message, "COVERED QTRS RESPONSE RECEIVED FROM SSA") then
    match_found = TRUE
    call run_from_GitHub(script_repository & "dail/tpqy-response.vbs")
END IF

'ABAWD Exemption DAIL Messages
If instr(full_message, "FSET WORK STATUS SHOWS UNDER AGE 16 FOR MEMBER 16 OR OLDER") or _
   instr(full_message, "SNAP ABAWD ELIGIBILITY HAS EXPIR") or _
   instr(full_message, "SNAP MEMBERS AGE IS 50 OR OLDER-REVIEW FOR ABAWD EXEMPTION") or _
   instr(full_message, "WREG PANEL INDICATES SNAP DISABILITY BUT DISA PANEL DOES NOT") or _
   instr(full_message, "WREG PANEL INDICATES UNEMP INS BUT UNEA PANEL DOES NOT") then
	match_found = TRUE
	CALL run_from_GitHub(script_repository & "dail/abawd-fset-exemption-check.vbs")
END IF

'UNBORN CHILD IS OVERDUE
EMReadScreen overdue_baby, 23, 6, 20
IF overdue_baby = "UNBORN CHILD IS OVERDUE" THEN
 	match_found = TRUE
	CALL run_from_GitHub(script_repository & "dail/overdue-baby.vbs")
END IF

IF match_found = FALSE THEN
    'WAGE MATCH Scrubber
    EMReadScreen DAIL_type, 4, 6, 6
    IF DAIL_type = "WAGE" THEN CALL run_from_GitHub(script_repository & "dail/wage-match-scrubber.vbs")

    'ALL other DAIL messages
    IF DAIL_type = "TIKL" or DAIL_type = "PEPR"  or DAIL_type = "INFO" THEN CALL run_from_GitHub(script_repository & "dail/catch-all.vbs")

    'Child support messages (loads CSES PROCESSING)
    IF DAIL_type = "CSES" THEN
    	EMReadScreen CSES_DISB_check, 4, 6, 20				'Checks for the DISB string, verifying this as a disbursement message
    	If CSES_DISB_check = "DISB" then call run_from_GitHub(script_repository & "dail/cses-scrubber.vbs") 'If it's a disbursement message...
    END IF
END IF

'NOW IF NO SCRIPT HAS BEEN WRITTEN FOR IT, THE DAIL SCRUBBER STOPS AND GENERATES A MESSAGE TO THE WORKER.----------------------------------------------------------------------------------------------------
script_end_procedure_with_error_report("You are not on a supported DAIL message. The script will now stop. " & vbNewLine & vbNewLine & "The message reads: " & full_message & vbNewLine & "Please send an error report if you would you like this DAIL to be supported.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/23/2022------------------N/A
'--Tab orders reviewed & confirmed----------------------------------------------05/23/2022------------------N/A
'--Mandatory fields all present & Reviewed--------------------------------------05/23/2022------------------N/A
'--All variables in dialog match mandatory fields-------------------------------05/23/2022------------------N/A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/23/2022------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------05/23/2022------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/23/2022------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/23/2022------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------05/23/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/23/2022
'--Out-of-County handling reviewed----------------------------------------------05/23/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/23/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/23/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/23/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/23/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/23/2022
'--Denomination reviewed -------------------------------------------------------05/23/2022
'--Script name reviewed---------------------------------------------------------05/23/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/23/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/23/2022
'--comment Code-----------------------------------------------------------------05/23/2022
'--Update Changelog for release/update------------------------------------------11/14/2022
'--Remove testing message boxes-------------------------------------------------05/23/2022
'--Remove testing code/unnecessary code-----------------------------------------05/23/2022
'--Review/update SharePoint instructions----------------------------------------11/14/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/23/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/23/2022
'--Complete misc. documentation (if applicable)---------------------------------05/23/2022
'--Update project team/issue contact (if applicable)----------------------------05/23/2022
