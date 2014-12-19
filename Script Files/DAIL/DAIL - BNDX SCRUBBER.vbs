'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - BNDX SCRUBBER.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

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

'========== Collects information from BNDX ==========
EMReadScreen bndx_claim_one_number, 14, 5, 12
  bndx_claim_one_number = replace(bndx_claim_one_number, " ", "")
EMReadScreen bndx_claim_one_amt, 8, 7, 12
  bndx_claim_one_amt = replace(bndx_claim_one_amt, " ", "")
  bndx_claim_one_amt = FormatCurrency(bndx_claim_one_amt)

EMReadScreen bndx_claim_two_number, 14, 5, 38
  bndx_claim_two_number = replace(bndx_claim_two_number, " ", "")
EMReadScreen bndx_claim_two_amt, 8, 7, 38
  bndx_claim_two_amt = replace(bndx_claim_two_amt, " ", "")
  IF bndx_claim_two_amt <> "" THEN bndx_claim_two_amt = INT(bndx_claim_two_amt)

EMReadScreen bndx_claim_three_number, 14, 5, 64
  bndx_claim_three_number = replace(bndx_claim_three_number, " ", "")
EMReadScreen bndx_claim_three_amt, 8, 7, 64
  bndx_claim_three_amt = replace(bndx_claim_three_amt, " ", "")
  IF bndx_claim_three_amt <> "" THEN bndx_claim_three_amt = INT(bndx_claim_three_amt)

'========== Goes back to STAT/PROG to determine which programs are active. ==========
back_to_SELF
EMWriteScreen "STAT", 16, 43
EMWriteScreen maxis_case_number, 18, 43
EMWriteScreen "PROG", 21, 70
transmit
abended_function
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


'========== Goes to STAT/UNEA ==========
EMWriteScreen "UNEA", 20, 71
EMWriteScreen reference_number, 20, 76
transmit

EMReadScreen number_of_unea_panels, 1, 2, 78
IF number_of_unea_panels = "1" THEN
	EMReadScreen unea_type, 4, 5, 40
	IF unea_type = "RSDI" THEN
		EMReadScreen unea_claim_number, 12, 6, 37
		IF right(unea_claim_number, 2) = "00" THEN unea_claim_number = left(unea_claim_number, 10)
		unea_claim_number = replace(unea_claim_number, "_", "")
		IF bndx_claim_one_number <> unea_claim_number THEN error_message = error_message & chr(13) & "Claim numbers do not match."
		EMReadScreen unea_prospective_amt, 8, 13, 68
		unea_prospective_amt = replace(unea_prospective_amt, " ", "")
		unea_prospective_amount = INT(unea_prospective_amt)
		IF ((unea_prospective_amount - bndx_claim_one_amt > county_bndx_variance_threshold) OR (bndx_claim_one_amt - unea_prospective_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amounts are significantly different."
		IF fs_status = "ACTV" or fs_status = "PEND" THEN
			EMWriteScreen "X", 10, 26
			transmit
			EMReadScreen unea_pic_amt, 8, 18, 56
			unea_pic_amt = replace(unea_pic_amt, " ", "")
			unea_pic_amount = INT(unea_pic_amt)
			IF ((unea_pic_amount - bndx_claim_one_amt > county_bndx_variance_threshold) OR (bndx_claim_one_amt - unea_pic_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX."
			PF3
		END IF
		IF hc_status = "ACTV" or hc_status = "PEND" THEN
			EMWriteScreen "X", 6, 56
			transmit
			EMReadScreen unea_hc_inc_amt, 8, 9, 65
			unea_hc_inc_amt = replace(unea_hc_inc_amt, " ", "")
			unea_hc_inc_amount = INT(unea_hc_inc_amt)
			IF ((unea_hc_inc_amount - bndx_claim_one_amt > county_bndx_variance_threshold) OR (bndx_claim_one_amt - unea_hc_inc_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the HC Income Estimator is significantly different from BNDX."
			PF3
		END IF
	ELSE
		error_message = "This case is not showing an RSDI claim."
	END IF
ELSEIF CINT(number_of_unea_panels) > 1 THEN
	number_of_unea_panels = cint(number_of_unea_panels)
	unea_panels_to_check = number_of_unea_panels
	claim_count = 0
	rsdi_count = 0
	'========== This part of the script collects information from UNEA when there are multiple UNEA screens. This is for display ONLY and is used 
	'========== to populate the MSGBox at the end of the script for Beta Testing. This DO-LOOP can be removed when the script is taken out of BETA.
	DO
		EMReadScreen unea_panel_number, 1, 2, 73
		unea_panel_number = cint(unea_panel_number)
		EMWriteScreen "0", 20, 79
		EMWriteScreen unea_panel_number, 20, 80
		transmit
		EMReadScreen unea_is_rsdi, 4, 5, 40
			IF unea_is_rsdi = "RSDI" THEN
				EMReadScreen unea_amount, 8, 13, 68
				unea_amount = replace(unea_amount, " ", "0")
				total_unea = unea_amount
				IF hc_status = "ACTV" OR hc_status = "PEND" THEN
					EMWriteScreen "X", 6, 56
					transmit
					EMReadScreen unea_hc_amount, 8, 9, 65
					unea_hc_amount = replace(unea_hc_amount, " ", "0")
					total_unea = total_unea & unea_hc_amount
					transmit
				END IF
				IF fs_status = "ACTV" OR fs_status = "PEND" THEN
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen uneaPIC, 8, 18, 56
					uneaPIC = replace(uneaPIC, " ", "0")
					total_unea = total_unea & uneaPIC
					transmit
				END IF
			END IF
		transmit
	LOOP UNTIL unea_panel_number = number_of_unea_panels
	DO
		EMReadScreen which_unea_panel, 1, 2, 73
		EMWriteScreen "0", 20, 79
		EMWriteScreen (claim_count + 1), 20, 80
		transmit
		EMReadScreen unea_type, 4, 5, 40
		IF unea_type = "RSDI" THEN
			EMReadScreen unea_claim_number, 10, 6, 37
			IF right(unea_claim_number, 2) = "00" THEN unea_claim_number = left(unea_claim_number, 10)
			unea_claim_number = replace(unea_claim_number, "_", "")
			IF bndx_claim_one_number = unea_claim_number THEN
				EMReadScreen unea_prospective_amt, 8, 13, 68
				unea_prospective_amt = replace(unea_prospective_amt, " ", "")
				unea_prospective_amount = INT(unea_prospective_amt)
				IF ((unea_prospective_amount - bndx_claim_one_amt > county_bndx_variance_threshold) OR (bndx_claim_one_amt - unea_prospective_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The amounts for claim 1 are significantly different."
				IF fs_status = "ACTV" or fs_status = "PEND" THEN
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen unea_pic_amt, 8, 18, 56
					unea_pic_amt = replace(unea_pic_amt, " ", "")
					unea_pic_amount = INT(unea_pic_amt)
					IF ((unea_pic_amount - bndx_claim_one_amt > county_bndx_variance_threshold) OR (bndx_claim_one_amt - unea_pic_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The amount on the PIC does not match the BNDX message for claim 1."
					PF3
				END IF
				IF hc_status = "ACTV" or hc_status = "PEND" THEN
					EMWriteScreen "X", 6, 56
					transmit
					EMReadScreen unea_hc_inc_amt, 8, 9, 65
					unea_hc_inc_amt = replace(unea_hc_inc_amt, " ", "")
					unea_hc_inc_amount = INT(unea_hc_inc_amt)
					IF ((unea_hc_inc_amount - bndx_claim_one_amt > county_bndx_variance_threshold) OR (bndx_claim_one_amt - unea_hc_inc_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the HC Income Estimator is significantly different from BNDX."
					PF3
				END IF
			ELSEIF bndx_claim_two_number = unea_claim_number THEN
				EMReadScreen unea_prospective_amt, 8, 13, 68
				unea_prospective_amt = replace(unea_prospective_amt, " ", "")
				unea_prospective_amount = INT(unea_prospective_amt)
				IF ((unea_prospective_amount - bndx_claim_two_amt > county_bndx_variance_threshold) OR (bndx_claim_two_amt - unea_prospective_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount for claim 2 is significantly different from BNDX."
				IF fs_status = "ACTV" or fs_status = "PEND" THEN
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen unea_pic_amt, 8, 18, 56
					unea_pic_amt = replace(unea_pic_amt, " ", "")
					unea_pic_amount = INT(unea_pic_amt)
					IF ((unea_pic_amount - bndx_claim_two_amt > county_bndx_variance_threshold) OR (bndx_claim_two_amt - unea_pic_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The amount on the PIC from claim 2 does not match the BNDX message."
					PF3
				END IF
				IF hc_status = "ACTV" or hc_status = "PEND" THEN
					EMWriteScreen "X", 6, 56
					transmit
					EMReadScreen unea_hc_inc_amt, 8, 9, 65
					unea_hc_inc_amt = replace(unea_hc_inc_amt, " ", "")
					unea_hc_inc_amount = INT(unea_hc_inc_amt)
					IF ((unea_hc_inc_amount - bndx_claim_two_amt > county_bndx_variance_threshold) OR (bndx_claim_two_amt - unea_hc_inc_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on claim 2 in the HC Income Estimator is significantly different from BNDX."
					PF3
				END IF
			ELSEIF bndx_claim_three_number = unea_claim_number THEN
				EMReadScreen unea_prospective_amt, 8, 13, 68
				unea_prospective_amt = replace(unea_prospective_amt, " ", "")
				unea_prospective_amount = INT(unea_prospective_amt)
				IF ((unea_prospective_amount - bndx_claim_three_amt > county_bndx_variance_threshold) OR (bndx_claim_three_amt - unea_prospective_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount for claim 3 is significantly different from BNDX."
				IF fs_status = "ACTV" or fs_status = "PEND" THEN
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen unea_pic_amt, 8, 18, 56
					unea_pic_amt = replace(unea_pic_amt, " ", "")
					unea_pic_amount = INT(unea_pic_amt)
					IF ((unea_pic_amount - bndx_claim_three_amt > county_bndx_variance_threshold) OR (bndx_claim_three_amt - unea_pic_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The amount on the PIC from claim 3 does not match the BNDX message."
					PF3
				END IF
				IF hc_status = "ACTV" or hc_status = "PEND" THEN
					EMWriteScreen "X", 6, 56
					transmit
					EMReadScreen unea_hc_inc_amt, 8, 9, 65
					unea_hc_inc_amt = replace(unea_hc_inc_amt, " ", "")
					unea_hc_inc_amount = INT(unea_hc_inc_amt)
					IF ((unea_hc_inc_amount - bndx_claim_three_amt > county_bndx_variance_threshold) OR (bndx_claim_three_amt - unea_hc_inc_amount > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on claim 3 in the HC Income Estimator is significantly different from BNDX."
					PF3
				END IF
			END IF
			rsdi_count = rsdi_count + 1
		END IF
		claim_count = claim_count + 1
		number_of_unea_panels = number_of_unea_panels + 1
	LOOP UNTIL claim_count = unea_panels_to_check
	IF rsdi_count = 0 THEN error_message = "This case is not showing any RSDI claims."
ELSEIF number_of_unea_panels = "0" THEN
	error_message = "This case is not showing any UNEA panels."
END IF

IF rsdi_count > 1 THEN error_message = "This client has multiple RSDI claims. Double check case and process." & chr(13) & error_message

back_to_SELF
EMWriteScreen "DAIL", 16, 43
EMWriteScreen maxis_case_number, 18, 43
EMWriteScreen "DAIL", 21, 70
transmit

'========== The bit about the MSGBox is used only as a safeguard for Beta Testing.
IF error_message = "" THEN 
	compare_message = "BNDX Claim 1 Amt: " & (bndx_claim_one_amt)
	IF bndx_claim_two_amt <> "" THEN compare_message = compare_message & chr(13) & "BNDX Claim 2 Amt: " & (bndx_claim_two_amt)
	IF bndx_claim_three_amt <> "" THEN compare_message = compare_message & chr(13) & "BNDX Claim 3 Amt: " & (bndx_claim_three_amt)
	IF total_unea <> "" THEN
		uneaPROS = left(total_unea, 8)
		compare_message = compare_message & chr(13) & "UNEA Prospected Amt: " & FormatCurrency(uneaPROS)
		IF hc_status = "ACTV" OR hc_status = "PEND" THEN
			uneaHCamt = right(left(total_unea, 16), 8)
			compare_message = compare_message & chr(13) & "UNEA HC Inc Est: " & FormatCurrency(uneaHCamt)
		END IF
		IF fs_status = "ACTV" OR fs_status = "PEND" THEN
			uneaFSPIC = right(total_unea, 8)
			compare_message = compare_message & chr(13) & "UNEA SNAP PIC: " & FormatCurrency(uneaFSPIC)
		END IF
	ELSE
		compare_message = compare_message & chr(13) & "UNEA Prospected Amt: " & FormatCurrency(unea_prospective_amt)
		IF hc_status = "ACTV" OR hc_status = "PEND" THEN 
			unea_hc_inc_amt = cstr(unea_hc_inc_amt)
			compare_message = compare_message & chr(13) & "UNEA HC Inc Est: " & FormatCurrency(unea_hc_inc_amt)
		END IF
		IF fs_status = "ACTV" OR fs_status = "PEND" THEN
			unea_pic_amt = cstr(unea_pic_amt)
			compare_message = compare_message & chr(13) & "UNEA SNAP PIC:   " & FormatCurrency(unea_pic_amt)
		END IF
	END IF
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
		ELSE
			script_end_procedure("Double check the case and try again. You may need to send an electronic SSA verif request.")
		END IF
ELSE
	MSGBox error_message & chr(13) & "Review case and request RSDI information if necessary."
	compare_message = "BNDX Claim 1 Amt: " & (bndx_claim_one_amt)
	IF bndx_claim_two_amt <> "" THEN compare_message = compare_message & chr(13) & "BNDX Claim 2 Amt: " & (bndx_claim_two_amt)
	IF bndx_claim_three_amt <> "" THEN compare_message = compare_message & chr(13) & "BNDX Claim 3 Amt: " & (bndx_claim_three_amt)
	IF total_unea <> "" THEN
		uneaPROS = left(total_unea, 8)
		compare_message = compare_message & chr(13) & "UNEA Prospected Amt: " & FormatCurrency(uneaPROS)
		IF hc_status = "ACTV" OR hc_status = "PEND" THEN
			uneaHCamt = right(left(total_unea, 16), 8)
			compare_message = compare_message & chr(13) & "UNEA HC Inc Est: " & FormatCurrency(uneaHCamt)
		END IF
		IF fs_status = "ACTV" OR fs_status = "PEND" THEN
			uneaFSPIC = right(total_unea, 8)
			compare_message = compare_message & chr(13) & "UNEA SNAP PIC: " & FormatCurrency(uneaFSPIC)
		END IF
	ELSE
		compare_message = compare_message & chr(13) & "UNEA Prospected Amt: " & FormatCurrency(unea_prospective_amt)
		IF hc_status = "ACTV" OR hc_status = "PEND" THEN 
			unea_hc_inc_amt = cstr(unea_hc_inc_amt)
			compare_message = compare_message & chr(13) & "UNEA HC Inc Est: " & FormatCurrency(unea_hc_inc_amt)
		END IF
		IF fs_status = "ACTV" OR fs_status = "PEND" THEN
			unea_pic_amt = cstr(unea_pic_amt)
			compare_message = compare_message & chr(13) & "UNEA SNAP PIC:   " & FormatCurrency(unea_pic_amt)
		END IF
	END IF
	MSGBox compare_message
END IF

script_end_procedure("")
