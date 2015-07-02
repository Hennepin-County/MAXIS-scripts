'This script has been updated on 04/23/2015.
'(1) This script has been updated to handle multiple RSDI claims. I have used a multidimensional array to handle a whole bunch of information. Here's how it goes...
'	The "i" in the bndx_array refers to the RSDI claim. If the client has 1 RSDI claim, it will work only with bndx_array(0, x). If there are two, it wll be bndx_array(1, x), etc.
'	bndx_array(i, 0) = RSDI claim number as found in BNDX
'	bndx_array(i, 1) = RSDI claim amount as found in BNDX
'	bndx_array(i, 2) = RSDI claim number as found in UNEA
'	bndx_array(i, 3) = RSDI prospective amount as found in UNEA
'	bndx_array(i, 4) = RSDI amount found in UNEA PIC
'	bndx_array(i, 5) = RSDI amount found in UNEA HC INC EST
'----------------------------------------------------------------
'(2) The error message and comparison message has been updated to provide more information.
'----------------------------------------------------------------
'(3) The way the script handles the RSDI claim number has been updated. Because some claim suffixes include numeric values, the script will read 11 characters on UNEA, but it will always drop the last 
' character when reading a claim number ending in "A". The reason for this is that some cases have A00 in the claim number and some only have "A". BNDX, however, will only ever put "A" for the suffix.
'-----------------------------------------------------------------
'(4) Future plans could include bulking the BNDX Scrubber. However, for 04/2015, the number of BNDX DAILs seems to be down considerably. That enhancement request could be put on hold.
'-----------------------------------------------------------------

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - BNDX SCRUBBER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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


FUNCTION abended_function
	EMReadScreen case_abended, 7, 9, 27
	IF case_abended = "abended" THEN transmit
END FUNCTION

BeginDialog delete_message_dialog, 0, 0, 126, 45, "Double-Check the Computer's Work..."
  ButtonGroup ButtonPressed
    PushButton 10, 25, 50, 15, "YES", delete_button
    PushButton 60, 25, 50, 15, "NO", do_not_delete
  Text 30, 10, 65, 10, "Delete the DAIL??"
EndDialog

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""

EMReadScreen on_dail, 4, 2, 48
IF on_dail <> "DAIL" THEN script_end_procedure("You are not in DAIL. Please navigate to DAIL and run the script again.")

EMGetCursor read_row, read_column

EMReadScreen is_right_line, 34, read_row, 30
IF is_right_line <> "BENDEX INFORMATION HAS BEEN STORED" THEN script_end_procedure("You are not on the correct line. Please select a BNDX message on your DAIL.")
EMReadScreen original_bndx_dail, 30, read_row, 6

EMReadScreen cl_ssn, 9, read_row, 20
	ssn_first = left(cl_ssn, 3)
	ssn_first = ssn_first & " "
	ssn_mid = right(left(cl_ssn, 5), 2)
	ssn_mid = ssn_mid & " "
	ssn_end = right(cl_ssn, 4)
	use_ssn = ssn_first & ssn_mid & ssn_end
search_row = read_row

'========== Collects the case number ==========
DO
	EMReadScreen look_for_case_number, 18, search_row, 63
	IF left(look_for_case_number, 10) = "CASE NBR: " THEN 
		maxis_case_number = right(look_for_case_number, 8)
		maxis_case_number = replace(maxis_case_number, " ", "")
	ELSE
		search_row = search_row - 1
	END IF
LOOP UNTIL left(look_for_case_number, 10) = "CASE NBR: "

EMWriteScreen "I", read_row, 3
transmit
EMWriteScreen "BNDX", 20, 71
transmit

'checking for IRS non-disclosure agreement.
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure(To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.)

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
'		ReDim bndx_array(2, 5)
		num_of_rsdi = 2
		bndx_array(2, 0) = bndx_claim_three_number
		bndx_array(2, 1) = bndx_claim_three_amt
	END IF


	
'========== Goes back to STAT/PROG to determine which programs are active. ==========
back_to_SELF
EMWriteScreen "STAT", 16, 43
EMWriteScreen maxis_case_number, 18, 43
EMWriteScreen "PROG", 21, 70
transmit
EMReadScreen abended_check, 7, 9, 27
IF abended_check = "abended" THEN transmit
EMReadScreen errr_check, 4, 2, 52
IF errr_check = "ERRR" THEN transmit

EMReadScreen cash_one_status, 4, 6, 74
EMReadScreen cash_two_status, 4, 7, 74
EMReadScreen grh_status, 4, 9, 74
EMReadScreen fs_status, 4, 10, 74
EMReadScreen ive_status, 4, 11, 74
EMReadScreen hc_status, 4, 12, 74

IF cash_one_status <> "ACTV" AND cash_two_status <> "ACTV" AND grh_status <> "ACTV" AND fs_status <> "ACTV" AND ive_status <> "ACTV" AND hc_status <> "ACTV" THEN 
  IF cash_one_status <> "PEND" AND cash_two_status <> "PEND" AND grh_status <> "PEND" AND fs_status <> "PEND" AND ive_status <> "PEND" AND hc_status <> "PEND" THEN script_end_procedure("The client does not have any active or pending MAXIS cases.")
END IF

'========== Navigates to MEMB to grab the member number for cases in which there are mulitple persons on the case with a BNDX message. ==========
EMWriteScreen "MEMB", 20, 71
transmit

DO
	EMReadScreen memb_ssn, 11, 7, 42
	IF use_ssn = memb_ssn THEN
		EMReadScreen reference_number, 2, 4, 33
	ELSE
		transmit
	END IF
LOOP UNTIL use_ssn = memb_ssn

FOR i = 0 TO num_of_rsdi
	end_of_unea = ""
	'========== Goes to STAT/UNEA ==========
	EMWriteScreen "UNEA", 20, 71
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "01", 20, 79
	transmit

	EMReadScreen number_of_unea_panels, 1, 2, 78
	IF number_of_unea_panels = "0" THEN
		script_end_procedure("Client is not showing any UNEA panels.")
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
			IF fs_status = "ACTV" or fs_status = "PEND" THEN
				EMWriteScreen "X", 10, 26
				transmit
				EMReadScreen unea_pic_amt, 8, 18, 56
				bndx_array(i, 4) = trim(unea_pic_amt)
				IF ((CDbl(bndx_array(i, 4)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 4)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
				PF3
			ELSE
				bndx_array(i, 4) = ""
			END IF
			IF hc_status = "ACTV" or hc_status = "PEND" THEN
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
				IF fs_status = "ACTV" or fs_status = "PEND" THEN
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen unea_pic_amt, 8, 18, 56
					bndx_array(i, 4) = trim(unea_pic_amt)
					IF ((CDbl(bndx_array(i, 4)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 4)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
					PF3
				ELSE
					bndx_array(i, 4) = ""
				END IF
				IF hc_status = "ACTV" or hc_status = "PEND" THEN
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
EMWriteScreen "DAIL", 16, 43
EMWriteScreen maxis_case_number, 18, 43
EMWriteScreen "DAIL", 21, 70
transmit

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
	DIALOG delete_message_dialog
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
	'compare_message = "BNDX Conclusion" & vbCr & "============="
	'FOR i = 0 to num_of_rsdi
	'	compare_message = compare_message & vbCr & "BNDX Claim #: " & bndx_array(i, 0)
	'	compare_message = compare_message & vbCr & "  BNDX Amt: " & bndx_array(i, 1) 
	'	compare_message = compare_message & vbCr & "  UNEA Prosp Amt: " & bndx_array(i, 3)
	'	IF bndx_array(i, 4) <> "" THEN compare_message = compare_message & vbCr & "  SNAP PIC Amt: " & bndx_array(i, 4)
	'	IF bndx_array(i, 5) <> "" THEN compare_message = compare_message & vbCr & "  HC Inc Est Amt: " & bndx_array(i, 5)
	'NEXT
	'MSGBox compare_message
END IF

script_end_procedure("")
