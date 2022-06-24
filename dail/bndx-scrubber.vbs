'Required for statistical purposes===============================================================================
name_of_script = "DAIL - BNDX SCRUBBER.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 80        'manual run time in seconds
STATS_denomination = "I"       'I is for each Bndx dail message
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
call changelog_update("06/21/2021", "Updated handling for non-disclosure agreement.", "MiKayla Handley") '#493
call changelog_update("05/16/2022", "Updated script functionality to support IEVS/INFO message updates. This DAIL scrubber will work on both older message with SSN's and new messages without.", "Ilse Ferris, Hennepin County") ''#814
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.
EMConnect ""

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)
MAXIS_case_number = replace(MAXIS_case_number, "_", "")

EmReadscreen MAXIS_footer_month, 2, 6, 11   'reading footer month in case there are more DAIL's then just the one for INFC.
EmReadscreen MAXIS_footer_year, 2, 6, 14

EMReadScreen MEMB_check, 7, 6, 20
If left(MEMB_check, 4) = "MEMB" then
    member_number = right(MEMB_check, 2)
    SSN_present = False
    'Grabbibng the SSN for the member
    EmReadscreen member_number, 2, 6, 25
    CALL write_value_and_transmit("S", 6, 3)
    'PRIV Handling
    EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
    If priv_check = "PRIVIL" THEN script_end_procedure("This case is priviledged. The script will now end.")
    EMWriteScreen "MEMB", 20, 71
    Call write_value_and_transmit(member_number, 20, 76)
    EmReadscreen client_SSN, 11, 7, 42
    trimmed_client_SSN = replace(client_SSN, " ", "")
    PF3 ' back to the DAIL
Else
    'DAIL messages that have the SSN already present don't need to enter the MEMB panel to gather the SSN
    SSN_present = True
    EMReadScreen cl_ssn, 9, 6, 20
    ssn_first = left(cl_ssn, 3)
    ssn_first = ssn_first & " "
    ssn_mid = right(left(cl_ssn, 5), 2)
    ssn_mid = ssn_mid & " "
    ssn_end = right(cl_ssn, 4)
    client_SSN = ssn_first & ssn_mid & ssn_end
End if

'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)
'PRIV Handling
EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
If priv_check = "PRIVIL" THEN script_end_procedure("This case is priviledged. The script will now end.")
EMWriteScreen trimmed_client_SSN, 3, 63

'Ensuring that we're on the right footer month/year as the DAIL message. Otherwise errors can occur.
EmReadscreen bndx_month, 2, 20, 55
EmReadscreen bndx_year, 2, 20, 58

If bndx_month <> MAXIS_footer_month then EmWriteScreen MAXIS_footer_month, 20, 55
If bndx_year <> MAXIS_footer_year then EmWriteScreen MAXIS_footer_year, 20, 58

CALL write_value_and_transmit("BNDX", 20, 71)
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

DIM bndx_array()
ReDim bndx_array(2, 5)
'========== Collects information from BNDX ==========
EMReadScreen bndx_claim_one_number, 13, 5, 12
bndx_claim_one_number = replace(bndx_claim_one_number, " ", "")
EMReadScreen bndx_claim_one_amt, 8, 7, 12
bndx_claim_one_amt = replace(bndx_claim_one_amt, " ", "")
'ReDim bndx_array(0, 5)
num_of_rsdi = 0
bndx_array(0, 0) = bndx_claim_one_number
bndx_array(0, 1) = bndx_claim_one_amt

EMReadScreen bndx_claim_two_number, 13, 5, 38
bndx_claim_two_number = replace(bndx_claim_two_number, " ", "")
EMReadScreen bndx_claim_two_amt, 8, 7, 38
bndx_claim_two_amt = replace(bndx_claim_two_amt, " ", "")
	IF bndx_claim_two_amt <> "" THEN
'		ReDim bndx_array(1, 5)
		num_of_rsdi = 1
		bndx_array(1, 0) = bndx_claim_two_number
		bndx_array(1, 1) = bndx_claim_two_amt
	END IF

EMReadScreen bndx_claim_three_number, 13, 5, 64
bndx_claim_three_number = replace(bndx_claim_three_number, " ", "")
EMReadScreen bndx_claim_three_amt, 8, 7, 64
bndx_claim_three_amt = replace(bndx_claim_three_amt, " ", "")
IF bndx_claim_three_amt <> "" THEN
'	ReDim bndx_array(2, 5)
	num_of_rsdi = 2
	bndx_array(2, 0) = bndx_claim_three_number
	bndx_array(2, 1) = bndx_claim_three_amt
END IF

PF3' back to DAIL

'========== Navigates to CASE/CURR from the DAIL to determine which programs are active/pending. ==========
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

If case_active = False then
    If case_pending = False then script_end_procedure("This case does not have any active or pending MAXIS programs.")
End if

PF3 'back to DAIL from CASE/CURR
'========== Navigates to MEMB to grab the member number for cases in which there are mulitple persons on the case with a BNDX message. ==========
Call write_value_and_transmit("S", 6, 3)
Call write_value_and_transmit("MEMB", 20, 71)

DO
	EMReadScreen memb_ssn, 11, 7, 42
	IF client_SSN = memb_ssn THEN
		EMReadScreen member_number, 2, 4, 33
	ELSE
		transmit
	END IF
LOOP UNTIL client_SSN = memb_ssn

FOR i = 0 TO num_of_rsdi
	end_of_unea = ""
	'========== Goes to STAT/UNEA ==========
	Call write_value_and_transmit("UNEA", 20, 71)
	EMWriteScreen member_number, 20, 76
	Call write_value_and_transmit("01", 20, 79)

	EMReadScreen number_of_unea_panels, 1, 2, 78
	IF number_of_unea_panels = "0" THEN
		script_end_procedure("Resident is not showing any UNEA panels.")
	ELSEIF number_of_unea_panels = "1" THEN
		EMReadScreen unea_type, 4, 5, 40
		IF unea_type = "RSDI" THEN
			EMReadScreen unea_claim_number, 11, 6, 37
			bndx_array(i, 2) = unea_claim_number
			IF (right(bndx_array(i, 0), 1) = "A" AND right(bndx_array(i, 0), 2) <> "HA") OR _
				(right(bndx_array(i, 0), 1) = "B" AND right(bndx_array(i, 0), 2) <> "HB") OR _
				right(bndx_array(i, 0), 1) = "D" OR _
				right(bndx_array(i, 0), 1) = "E" OR _
				right(bndx_array(i, 0), 1) = "G" OR _
				right(bndx_array(i, 0), 1) = "M" OR _
				right(bndx_array(i, 0), 1) = "T" OR _
				right(bndx_array(i, 0), 1) = "W" THEN bndx_array(i, 2) = left(bndx_array(i, 2), 10)

			IF bndx_array(i, 0) <> bndx_array(i, 2) THEN error_message = error_message & chr(13) & "Claim numbers do not match."
			EMReadScreen unea_prospective_amt, 8, 18, 68
			bndx_array(i, 3) = trim(unea_prospective_amt)
			IF ((CDbl(bndx_array(i, 3)) - CDBl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 3)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The prospective amount in UNEA is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
			IF snap_case = True then
				EMWriteScreen "X", 10, 26
				transmit
				EMReadScreen unea_pic_amt, 8, 18, 56
				bndx_array(i, 4) = trim(unea_pic_amt)
				IF ((CDbl(bndx_array(i, 4)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 4)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
				PF3
			ELSE
				bndx_array(i, 4) = ""
			END IF
			IF (ma_case = True or msa_case = True or unknown_hc_pending = True) then
				EMWriteScreen "X", 6, 56
				transmit
				EMReadScreen unea_hc_inc_amt, 8, 9, 65
				bndx_array(i, 5) = trim(unea_hc_inc_amt)
				IF ((CDbl(bndx_array(i, 5)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 5)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the HC Inc Est is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
				PF3
			ELSE
				bndx_array(i, 5) = ""
			END IF
		ELSE
			script_end_procedure("This case is not showing an RSDI claim.")
		END IF
	ELSE
		DO
			EMReadScreen unea_type, 4, 5, 40
			IF unea_type <> "RSDI" THEN transmit
			EMReadScreen end_of_unea, 15, 24, 2
			end_of_unea = trim(end_of_unea)
			IF end_of_unea <> "" THEN error_message = error_message & vbCr & "There is a discrepancy with BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
		LOOP UNTIL unea_type = "RSDI" or end_of_unea <> ""
		IF end_of_unea = "" THEN
			DO
				EMReadScreen unea_claim_number, 11, 6, 37
				bndx_array(i, 2) = unea_claim_number
				IF (right(bndx_array(i, 0), 1) = "A" AND right(bndx_array(i, 0), 2) <> "HA") OR _
					(right(bndx_array(i, 0), 1) = "B" AND right(bndx_array(i, 0), 2) <> "HB") OR _
					right(bndx_array(i, 0), 1) = "D" OR _
					right(bndx_array(i, 0), 1) = "E" OR _
					right(bndx_array(i, 0), 1) = "G" OR _
					right(bndx_array(i, 0), 1) = "M" OR _
					right(bndx_array(i, 0), 1) = "T" OR _
					right(bndx_array(i, 0), 1) = "W" THEN bndx_array(i, 2) = left(bndx_array(i, 2), 10)
				IF bndx_array(i, 0) <> bndx_array(i, 2) THEN transmit
				EMReadScreen end_of_unea, 15, 24, 2
				end_of_unea = trim(end_of_unea)
				IF end_of_unea <> "" THEN error_message = error_message & vbCr & "There is a discrepancy with BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
			LOOP UNTIL bndx_array(i, 0) = bndx_array(i, 2) OR end_of_unea <> ""
			IF end_of_unea = "" THEN
				EMReadScreen unea_prospective_amt, 8, 18, 68
				bndx_array(i, 3) = trim(unea_prospective_amt)
				IF ((CDbl(bndx_array(i, 3)) - CDBl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 3)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The prospective amount in UNEA is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
				IF snap_case = True then
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen unea_pic_amt, 8, 18, 56
					bndx_array(i, 4) = trim(unea_pic_amt)
					IF ((CDbl(bndx_array(i, 4)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 4)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
					PF3
				ELSE
					bndx_array(i, 4) = ""
				END IF
				IF (ma_case = True or msa_case = True or unknown_hc_pending = True) then
					EMWriteScreen "X", 6, 56
					transmit
					EMReadScreen unea_hc_inc_amt, 8, 9, 65
					bndx_array(i, 5) = trim(unea_hc_inc_amt)
					IF ((CDbl(bndx_array(i, 5)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDBl(bndx_array(i, 5)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the HC Inc Est is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
					PF3
				ELSE
					bndx_array(i, 5) = ""
				END IF
			END IF
		END IF
	END IF
NEXT

back_to_SELF
Call navigate_to_MAXIS_screen("DAIL", "DAIL")

'========== The bit about the MSGBox is used only as a safeguard for Beta Testing.
IF error_message = "" THEN
	compare_message = "BNDX Conclusion" & vbCr & "============="
	FOR i = 0 to num_of_rsdi
		compare_message = compare_message & vbCr & "BNDX Claim #: " & bndx_array(i, 0)
		compare_message = compare_message & vbCr & "  BNDX Amt: " & bndx_array(i, 1)
		compare_message = compare_message & vbCr & "  UNEA Prosp Amt: " & bndx_array(i, 3)
		IF bndx_array(i, 4) <> "" THEN compare_message = compare_message & vbCr & "  SNAP PIC Amt: " & bndx_array(i, 4)
		IF bndx_array(i, 5) <> "" THEN compare_message = compare_message & vbCr & "  HC Inc Est Amt: " & bndx_array(i, 5)
	NEXT
	MSGBox compare_message
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 126, 45, "Double-Check the Computer's Work..."
      ButtonGroup ButtonPressed
        PushButton 10, 25, 50, 15, "YES", delete_button
        PushButton 60, 25, 50, 15, "NO", do_not_delete
      Text 30, 10, 65, 10, "Delete the DAIL??"
    EndDialog

DIALOG Dialog1
		IF ButtonPressed = delete_button THEN
			DO
				dail_read_row = 6
				DO
					EMReadScreen double_check, 30, dail_read_row, 6
					IF double_check = original_bndx_dail THEN
						EMWriteScreen "D", dail_read_row, 3
						transmit
						EXIT DO
					ELSE
						dail_read_row = dail_read_row + 1
					END IF
					IF dail_read_row = 19 THEN PF8
				LOOP UNTIL dail_read_row = 19
			LOOP UNTIL double_check = original_bndx_dail
		END IF
ELSE
	error_message = "*** NOTICE ***" & vbCr & "==========" & vbCr & error_message & vbCr & vbCr & "Review case and request RSDI information if necessary."
	MSGBox error_message
END IF

script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/25/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/25/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/25/2022
'--All variables in dialog match mandatory fields-------------------------------05/25/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/25/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------05/25/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/25/2022-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/25/2022-------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------05/25/2022-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/25/2022
'--Out-of-County handling reviewed----------------------------------------------05/25/2022-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/25/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/25/2022-------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/25/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/25/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/25/2022
'--Denomination reviewed -------------------------------------------------------05/25/2022
'--Script name reviewed---------------------------------------------------------05/25/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/25/2022-------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/25/2022
'--comment Code-----------------------------------------------------------------05/25/2022
'--Update Changelog for release/update------------------------------------------05/16/2022
'--Remove testing message boxes-------------------------------------------------05/25/2022
'--Remove testing code/unnecessary code-----------------------------------------05/25/2022
'--Review/update SharePoint instructions----------------------------------------05/25/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/25/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/25/2022
'--Complete misc. documentation (if applicable)---------------------------------05/25/2022
'--Update project team/issue contact (if applicable)----------------------------05/25/2022
