'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Project Krabappel\KRABAPPEL FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

hc_application = True
case_number = "213161"
appl_date_month = 11
appl_date_year = 14

EMConnect ""

	IF hc_application = TRUE THEN
		call navigate_to_screen("ELIG", "HC")
		'======The script needs to select a HH member to approve HC

		'=====This part of the script makes the FIAT changes to HH members with Budg Mthd A
		hhmm_row = 8
		DO
			EMReadScreen hc_requested, 1, hhmm_row, 28
			EMReadScreen hc_status, 5, hhmm_row, 68
			IF hc_requested = "M" AND hc_status = "UNAPP" THEN 
				DO						'===== This DO/LOOP is for the check to determine the case is not stuck in ELIG. If it is, it will not let you FIAT Elig Standard.
					EMWriteScreen "X", hhmm_row, 26
					transmit				'===== Navigates to BSUM for the HH member
	
					EMReadScreen budg_mthd_mo1, 1, 13, 21
					EMReadScreen budg_mthd_mo2, 1, 13, 32
					EMReadScreen budg_mthd_mo3, 1, 13, 43
					EMReadScreen budg_mthd_mo4, 1, 13, 54
					EMReadScreen budg_mthd_mo5, 1, 13, 65
					EMReadScreen budg_mthd_mo6, 1, 13, 76
					
					IF (budg_mthd_mo1 = "B") AND (budg_mthd_mo2 = "B") AND (budg_mthd_mo3 = "B") AND (budg_mthd_mo4 = "B") AND (budg_mthd_mo5 = "B") AND (budg_mthd_mo6 = "B") THEN
						PF3
						EXIT DO
					ELSEIF (budg_mthd_mo1 = "A") OR (budg_mthd_mo2 = "A") OR (budg_mthd_mo3 = "A") OR (budg_mthd_mo4 = "A") OR (budg_mthd_mo5 = "A") OR (budg_mthd_mo6 = "A") THEN
							PF9
							DO
								EMReadScreen fiat_reason, 4, 10, 20		'=====The script gets stuck in ELIG background...it's running faster than the training region will allow.
							LOOP UNTIL fiat_reason = "FIAT"
							EMWriteScreen "05", 11, 26
							transmit
						IF budg_mthd_mo1 = "A" THEN
							EMWriteScreen "X", 7, 17
							EMWriteScreen "X", 9, 21
							EMReadScreen mo1_elig_type, 2, 12, 17
							IF (mo1_elig_type = "AX" OR mo1_elig_type = "AA" OR mo1_elig_type = "CX") THEN EMWriteScreen "J", 12, 22
							IF (mo1_elig_type = "PX" OR mo1_elig_type = "PC") THEN EMWriteScreen "T", 12, 22
							IF (mo1_elig_type = "CK") THEN EMWriteScreen "K", 12, 22
							IF mo1_elig_type = "CB" THEN EMWriteScreen "I", 12, 22
						END IF
						IF budg_mthd_mo2 = "A" THEN
							EMWriteScreen "X", 7, 28
							EMWriteScreen "X", 9, 32
							EMReadScreen mo2_elig_type, 2, 12, 28
							IF (mo2_elig_type = "AX" OR mo2_elig_type = "AA" OR mo2_elig_type = "CX") THEN EMWriteScreen "J", 12, 33
							IF (mo2_elig_type = "PX" OR mo2_elig_type = "PC") THEN EMWriteScreen "T", 12, 33
							IF (mo2_elig_type = "CK") THEN EMWriteScreen "K", 12, 33
							IF mo2_elig_type = "CB" THEN EMWriteScreen "I", 12, 33		
						END IF
						IF budg_mthd_mo3 = "A" THEN
							EMWriteScreen "X", 7, 39
							EMWriteScreen "X", 9, 43
							EMReadScreen mo3_elig_type, 2, 12, 39
							IF (mo3_elig_type = "AX" OR mo3_elig_type = "AA" OR mo3_elig_type = "CX") THEN EMWriteScreen "J", 12, 44
							IF (mo3_elig_type = "PX" OR mo3_elig_type = "PC") THEN EMWriteScreen "T", 12, 44
							IF (mo3_elig_type = "CK") THEN EMWriteScreen "K", 12, 44
							IF mo3_elig_type = "CB" THEN EMWriteScreen "I", 12, 44						
						END IF
						IF budg_mthd_mo4 = "A" THEN
							EMWriteScreen "X", 7, 50
							EMWriteScreen "X", 9, 54
							EMReadScreen mo4_elig_type, 2, 12, 50
							IF (mo4_elig_type = "AX" OR mo4_elig_type = "AA" OR mo4_elig_type = "CX") THEN EMWriteScreen "J", 12, 55
							IF (mo4_elig_type = "PX" OR mo4_elig_type = "PC") THEN EMWriteScreen "T", 12, 55
							IF (mo4_elig_type = "CK") THEN EMWriteScreen "K", 12, 55
							IF mo4_elig_type = "CB" THEN EMWriteScreen "I", 12, 55			
						END IF
						IF budg_mthd_mo5 = "A" THEN
							EMWriteScreen "X", 7, 61
							EMWriteScreen "X", 9, 65
							EMReadScreen mo5_elig_type, 2, 12, 61
							IF (mo5_elig_type = "AX" OR mo5_elig_type = "AA" OR mo5_elig_type = "CX") THEN EMWriteScreen "J", 12, 66
							IF (mo5_elig_type = "PX" OR mo5_elig_type = "PC") THEN EMWriteScreen "T", 12, 66
							IF (mo5_elig_type = "CK") THEN EMWriteScreen "K", 12, 66
							IF mo5_elig_type = "CB" THEN EMWriteScreen "I", 12, 66			
						END IF
						IF budg_mthd_mo6 = "A" THEN
							EMWriteScreen "X", 7, 72
							EMWriteScreen "X", 9, 76
							EMReadScreen mo6_elig_type, 2, 12, 72
							IF (mo6_elig_type = "AX" OR mo6_elig_type = "AA" OR mo6_elig_type = "CX") THEN EMWriteScreen "J", 12, 77
							IF (mo6_elig_type = "PX" OR mo6_elig_type = "PC") THEN EMWriteScreen "T", 12, 77
							IF (mo6_elig_type = "CK") THEN EMWriteScreen "K", 12, 77
							IF mo6_elig_type = "CB" THEN EMWriteScreen "I", 12, 77			
						END IF
						transmit		'IF Budg Mthd A, transmit to navigate to MAPT & CBUD
						DO
							EMReadScreen back_to_bsum, 4, 3, 57
							IF back_to_BSUM <> "BSUM" THEN 
								EMReadScreen mapt, 4, 3, 51
								EMReadScreen cbud, 4, 3, 54
								EMReadScreen abud, 4, 3, 47
								IF mapt = "MAPT" THEN
									EMWriteScreen "PASSED", 8, 46		'=====Passes MNSure test
									transmit
									transmit
								END IF
								IF cbud = "CBUD" OR abud = "ABUD" THEN transmit
							END IF
						LOOP UNTIL back_to_bsum = "BSUM"

					END IF

					EMReadScreen cannot_fiat, 10, 24, 6
					IF cannot_fiat <> "          " THEN 		'===== IF the case is stuck in ELIG, it will not allow you to change the ELIG standard to the ACA-appropriate standard.
						PF10							'===== The script OOPS's the FIAT and backs out. It will reread and re-transmit the FIAT'd elig information.
						PF3
						EMWriteScreen "WAIT", 20, 71
						EMWaitReady 2, 2000
						EMWriteScreen "____", 20, 71
					ELSE
						PF3
					END IF
					
				LOOP UNTIL cannot_fiat = "          "

			END IF

			hhmm_row = hhmm_row + 1

		LOOP UNTIL hc_requested = " "			'===== Loops until there are no more HC versions to review

		'===== Now the script goes back in and approves everything.
		hhmm_row = 8
		DO
			EMReadScreen hc_requested, 1, hhmm_row, 28
			EMReadScreen hc_status, 5, hhmm_row, 68
			IF (hc_requested = "M" OR hc_requested = "S" OR hc_requested = "Q") AND hc_status = "UNAPP" THEN 
				EMWriteScreen "X", hhmm_row, 26
				transmit
				DO				
					EMReadScreen bhsm, 4, 3, 55
					EMReadScreen mesm, 4, 3, 56
					IF hc_requested = "M" THEN
						IF bhsm <> "BHSM" THEN
							transmit
						END IF
					ELSEIF hc_requested = "S" OR hc_requested = "Q" THEN
						IF mesm <> "MESM" THEN
							transmit
						END IF
					END IF
				LOOP UNTIL bhsm = "BHSM" OR mesm = "MESM"

				EMWriteScreen "APP", 20, 71
				transmit
 	
				'=====This portion of the script selects the possible HC programs and places an X on all of them for approval.=====
				FOR i = 9 to 24
					EMReadScreen hc_program, 1, i, 5
					IF hc_program = "_" THEN EMWriteScreen "X", i, 5
				NEXT
				transmit
		
				DO
					'=====This checks for a PRISM referral
					EMReadScreen prism_referral, 5, 6, 27
					IF prism_referral = "PRISM" THEN 
						EMWriteScreen "N", 15, 47
						transmit
					END IF
				LOOP UNTIL prism_referral <> "PRISM"

				DO
					EMReadScreen continue_yn, 8, 21, 30
					IF continue_yn = "Continue" THEN
						EMWriteScreen "Y", 21, 46
						transmit
					END IF
				LOOP UNTIL continue_yn = "Continue"
			END IF
			hhmm_row = hhmm_row + 1
		LOOP UNTIL hc_requested = " "			'===== Loops until there are no more HC versions to review

	END IF					'=====THE END=====
