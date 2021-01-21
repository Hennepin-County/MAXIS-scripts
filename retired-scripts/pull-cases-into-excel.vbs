'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - PULL CASES INTO EXCEL.vbs"
start_time = timer
Stats_counter = 1
STATS_manualtime = 0            'manual run time in seconds
STATS_denomination = "C"       'C is for each Case

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
call changelog_update("08/07/2019", "Updated coding to read citizenship verif code at new location due to MEMI panel changes associated with New Spouse Income Policy.", "Ilse Ferris, Hennepin County")
call changelog_update("08/06/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------FUNCTIONS----------
FUNCTION check_panels_function(x, panel_status)
	FOR EACH hh_person IN x
		CALL navigate_to_MAXIS_screen("STAT", "HEST")
			IF hh_person <> "01" THEN
				EMWriteScreen hh_person, 20, 76
				transmit
			END IF
			EMReadScreen hest_info, 1, 2, 73
			IF hest_info <> "0" THEN panel_status = panel_status & "HEST"
			IF hest_info = "0" THEN
				EMWriteScreen "SHEL", 20, 71
				transmit
				IF hh_person <> "01" THEN
					EMWriteScreen hh_person, 20, 76
					transmit
				END IF
				EMReadScreen shel_info, 1, 2, 73
				IF shel_info <> "0" THEN panel_status = "SHEL"
				IF shel_info = "0" THEN
					EMWriteScreen "COEX", 20, 71
					transmit
					IF hh_person <> "01" THEN
						EMWriteScreen hh_person, 20, 76
						transmit
					END IF
					EMReadScreen coex_info, 1, 2, 73
					IF coex_info <> "0" THEN panel_status = "COEX"
					IF coex_info = "0" THEN
						EMWriteScreen "DCEX", 20, 71
						transmit
						IF hh_person <> "01" THEN
							EMWriteScreen hh_person, 20, 76
							transmit
						END IF
						EMReadScreen dcex_info, 1, 2, 73
						IF dcex_info = "0" THEN
							EMWriteScreen "BUSI", 20, 71
							transmit
							IF hh_person <> "01" THEN
								EMWriteScreen hh_person, 20, 76
								transmit
							END IF
							EMReadScreen busi_info, 1, 2, 73
							EMReadScreen busi_end_date, 8, 5, 71
							IF busi_info = "0" OR busi_end_date <> "__ __ __" THEN
								EMWriteScreen "UNEA", 20, 71
								transmit
								IF hh_person <> "01" THEN
									EMWriteScreen hh_person, 20, 76
									transmit
								END IF
								EMReadScreen unea_info, 1, 2, 73
								IF unea_info <> "0" THEN
									DO
										EMReadScreen unea_end_date, 8, 9, 68
										IF unea_end_date <> "__ __ __" THEN
											transmit
											EMReadScreen valid_command, 21, 24, 2
										END IF
									LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR unea_end_date = "__ __ __"
								END IF
								IF (unea_info <> "0" AND unea_end_date = "__ __ __") THEN panel_status = "UNEA"
								IF (unea_info <> "0" AND valid_command = "ENTER A VALID COMMAND") OR unea_info = "0" THEN
									CALL navigate_to_MAXIS_screen("STAT", "JOBS")
									IF hh_person <> "01" THEN
										EMWriteScreen hh_person, 20, 76
										transmit
									END IF
									EMReadScreen jobs_info, 1, 2, 73
									IF jobs_info <> "0" THEN
										DO
											EMReadScreen jobs_end_date, 8, 9, 49
											IF jobs_end_date <> "__ __ __" THEN
												transmit
												EMReadScreen valid_command, 21, 24, 2
											END IF
										LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR jobs_end_date = "__ __ __"
									END IF
									IF (jobs_info <> "0" AND jobs_end_date = "__ __ __") THEN panel_status = "JOBS"
								END IF
							END IF
						END IF
					END IF
				END IF
			END IF
	NEXT
END FUNCTION

FUNCTION check_panels_for_income(pers_array, panel_status)
	FOR EACH hh_person IN pers_array
		IF hh_person <> "" THEN
			'Checking BUSI
			CALL navigate_to_MAXIS_screen("STAT", "BUSI")
			IF hh_person <> "01" THEN
				EMWriteScreen hh_person, 20, 76
				transmit
			END IF
			EMReadScreen busi_info, 1, 2, 73
			EMReadScreen busi_end_date, 8, 5, 71
			IF busi_info = "0" OR busi_end_date <> "__ __ __" THEN
				'Checking UNEA
				EMWriteScreen "UNEA", 20, 71
				transmit
				IF hh_person <> "01" THEN
					EMWriteScreen hh_person, 20, 76
					transmit
				END IF
				EMReadScreen unea_info, 1, 2, 73
				IF unea_info <> "0" THEN
					DO
						EMReadScreen unea_end_date, 8, 9, 68
						IF unea_end_date <> "__ __ __" THEN
							transmit
							EMReadScreen valid_command, 21, 24, 2
						END IF
					LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR unea_end_date = "__ __ __"
				END IF
				IF (unea_info <> "0" AND unea_end_date = "__ __ __") THEN panel_status = "UNEA"
				IF (unea_info <> "0" AND valid_command = "ENTER A VALID COMMAND") OR unea_info = "0" THEN
					'Checking JOBS
					CALL navigate_to_MAXIS_screen("STAT", "JOBS")
					IF hh_person <> "01" THEN
						EMWriteScreen hh_person, 20, 76
						transmit
					END IF
					EMReadScreen jobs_info, 1, 2, 73
					IF jobs_info <> "0" THEN
						DO
							EMReadScreen jobs_end_date, 8, 9, 49
							IF jobs_end_date <> "__ __ __" THEN
								transmit
								EMReadScreen valid_command, 21, 24, 2
							END IF
						LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR jobs_end_date = "__ __ __"
					END IF
					IF (jobs_info <> "0" AND jobs_end_date = "__ __ __") THEN panel_status = "JOBS"
				END IF
			ELSE
				panel_status = "BUSI"
			END IF
		END IF
	NEXT
END FUNCTION

FUNCTION check_panels_for_hc_function(x, panel_status)
	FOR EACH hh_person IN x
		IF hh_person <> "" THEN
			' >>>>> UNEA <<<<<
			CALL navigate_to_MAXIS_screen("STAT", "UNEA")
				IF hh_person <> "01" THEN
					EMWriteScreen hh_person, 20, 76
					transmit
				END IF
				EMReadScreen unea_info, 1, 2, 73
				IF unea_info <> "0" THEN
					DO
						EMReadScreen unea_end_date, 8, 9, 68
						IF unea_end_date <> "__ __ __" THEN
							transmit
							EMReadScreen valid_command, 21, 24, 2
						END IF
					LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR unea_end_date = "__ __ __"
				END IF
				IF (unea_info <> "0" AND unea_end_date = "__ __ __") THEN panel_status = "UNEA"
				IF (unea_info <> "0" AND valid_command = "ENTER A VALID COMMAND") OR unea_info = "0" THEN
			' >>>>> JOBS <<<<<
					CALL navigate_to_MAXIS_screen("STAT", "JOBS")
					IF hh_person <> "01" THEN
						EMWriteScreen hh_person, 20, 76
						transmit
					END IF
					EMReadScreen jobs_info, 1, 2, 73
					IF jobs_info <> "0" THEN
						DO
							EMReadScreen jobs_end_date, 8, 9, 49
							IF jobs_end_date <> "__ __ __" THEN
								transmit
								EMReadScreen valid_command, 21, 24, 2
							END IF
						LOOP UNTIL valid_command = "ENTER A VALID COMMAND" OR jobs_end_date = "__ __ __"
					END IF
					IF (jobs_info <> "0" AND jobs_end_date = "__ __ __") THEN
						panel_status = "JOBS"
					ELSE
			' >>>>> ACCT <<<<<
						EMWriteScreen "ACCT", 20, 71
						IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
						transmit
						EMReadScreen acct_info, 1, 2, 73
						IF acct_info <> "0" THEN
							DO
								EMReadScreen valid_command, 21, 24, 2
								IF valid_command <> "ENTER A VALID COMMAND" THEN
									EMReadScreen counted_for_hc, 1, 14, 64
									IF counted_for_hc = "Y" THEN
										panel_status = "ACCT"
										EXIT DO
									ELSE
										transmit
									END IF
								END IF
							LOOP UNTIL valid_command = "ENTER A VALID COMMAND"
							'IF counted_for_hc = "Y" THEN EXIT FOR
						END IF
						IF acct_info = "0" OR (acct_info <> "0" AND counted_for_hc <> "Y") THEN
			' >>>>> CASH <<<<<
							EMWriteScreen "CASH", 20, 71
							IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
							transmit
							EMReadScreen cash_info, 1, 2, 73
							IF cash_info <> "0" THEN
								panel_status = "CASH"
							ELSE
			' >>>>> CARS <<<<<
								EMWriteScreen "CARS", 20, 71
								IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
								transmit
								EMReadScreen cars_info, 1, 2, 73
								IF cars_info <> "0" THEN
									panel_status = "CARS"
								ELSE
			' >>>>> SECU <<<<<
									EMWriteScreen "SECU", 20, 71
									IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
									transmit
									EMReadScreen secu_info, 1, 2, 73
									IF secu_info <> "0" THEN
										DO
											EMReadScreen valid_command, 21, 24, 2
											IF valid_command <> "ENTER A VALID COMMAND" THEN
												EMReadScreen counted_for_hc, 1, 15, 64
												IF counted_for_hc = "Y" THEN
													panel_status = "SECU"
													EXIT DO
												ELSE
													transmit
												END IF
											END IF
										LOOP UNTIL valid_command = "ENTER A VALID COMMAND"
										'IF counted_for_hc = "Y" THEN EXIT FOR
									END IF
									IF secu_info = "0" OR (secu_info <> "0" AND counted_for_hc <> "Y") THEN
			' >>>>> OTHR <<<<<
										EMWriteScreen "OTHR", 20, 71
										IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
										transmit
										EMReadScreen othr_info, 1, 2, 73
										IF othr_info <> "0" THEN
											DO
												EMReadScreen valid_command, 21, 24, 2
												IF valid_command <> "ENTER A VALID COMMAND" THEN
													EMReadScreen counted_for_hc, 1, 12, 64
													IF counted_for_hc = "Y" THEN
														panel_status = "OTHR"
														EXIT DO
													ELSE
														transmit
													END IF
												END IF
											LOOP UNTIL valid_command = "ENTER A VALID COMMAND"
											'IF counted_for_hc = "Y" THEN EXIT FOR
										END IF
										IF othr_info = "0" OR (othr_info <> "0" AND counted_for_hc <> "Y") THEN
			' >>>>> BUSI <<<<<
											EMWriteScreen "BUSI", 20, 71
											IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
											transmit
											EMReadScreen busi_info, 1, 2, 73
											IF busi_info <> "0" THEN
												DO
													EMReadScreen valid_command, 21, 24, 2
													IF valid_command <> "ENTER A VALID COMMAND" THEN
														EMReadScreen BUSI_end_date, 8, 5, 72
														IF BUSI_end_date = "__ __ __" THEN
															panel_status = "BUSI"
															EXIT DO
														ELSE
															transmit
														END IF
													END IF
												LOOP UNTIL valid_command = "ENTER A VALID COMMAND"
												'IF BUSI_end_date = "__ __ __" THEN EXIT FOR
											END IF
											IF busi_info = "0" OR (busi_info <> "0" AND BUSI_end_date <> "__ __ __") THEN
			' >>>>> SPON <<<<<
												EMWriteScreen "SPON", 20, 71
												IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
												transmit
												EMReadScreen spon_info, 1, 2, 73
												IF spon_info <> "0" THEN
													EMReadScreen counted_for_hc, 30, 15, 50
													counted_for_hc = replace(counted_for_hc, "$", "")
													counted_for_hc = replace(counted_for_hc, " ", "")
													counted_for_hc = replace(counted_for_hc, "_", "")
													IF counted_for_hc <> "" THEN panel_status = "SPON"
												END IF
												IF spon_info = "0" OR (spon_info <> "0" AND counted_for_hc = "") THEN
			' >>>>> REST <<<<<
													EMWriteScreen "REST", 20, 71
													IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
													transmit
													EMReadScreen rest_info, 1, 2, 73
													IF rest_info <> "0" THEN
														panel_status = "REST"
													ELSE
			' >>>>> RBIC <<<<<
														EMWriteScreen "RBIC", 20, 71
														IF hh_person <> "01" THEN EMWriteScreen hh_person, 20, 76
														transmit
														EMReadScreen rbic_info, 1, 2, 73
														IF rbic_info <> "0" THEN
															DO
																EMReadScreen valid_command, 21, 24, 2
																IF valid_command <> "ENTER A VALID COMMAND" THEN
																	EMReadScreen RBIC_end_date, 8, 6, 68
																	IF RBIC_end_date = "__ __ __" THEN
																		panel_status = "RBIC"
																		EXIT DO
																	ELSE
																		transmit
																	END IF
																END IF
															LOOP UNTIL valid_command = "ENTER A VALID COMMAND"
														END IF
													END IF
												END IF
											END IF
										END IF
									END IF
								END IF
							END IF
						END IF
					END IF
				END IF
		END IF
	NEXT
END FUNCTION

FUNCTION navigate_to_MMIS
	attn

	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	'The following will select the correct version of MMIS. First it looks for C327, then EK01, then C427.
	row = 1
	col = 1
	EMSearch "C327", row, col
	If row <> 0 then
		If row <> 1 then 'It has to do this in case the worker only has one option (as many LTC and OSA workers don't have the option to decide between MAXIS and MCRE case access). The MMIS screen will show the text, but it's in the first row in these instances.
			EMWriteScreen "x", row, 4
			transmit
		End if
	Else 'Some staff may only have EK01 (MMIS MCRE). The script will allow workers to use that if applicable.
		row = 1
		col = 1
		EMSearch "EK01", row, col
		If row <> 0 then
			If row <> 1 then
				EMWriteScreen "x", row, 4
				transmit
			End if
		Else 'Some OSAs have C427 (limited access). This will search for that.
			row = 1
			col = 1
			EMSearch "C427", row, col
			If row <> 0 then
				If row <> 1 then
					EMWriteScreen "x", row, 4
					transmit
				End if
			Else 'Some OSAs have EKIQ (limited MCRE access). This will search for that.
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 then
					If row <> 1 then
						EMWriteScreen "x", row, 4
						transmit
					End if
				Else
					script_end_procedure("C427, C327, EKIQ, or EK01 not found. Your access to MMIS may be limited. Contact your script Alpha user if you have questions about using this script.")
				End if
			End if
		End if
	END IF

	'Now it finds the recipient file application feature and selects it.
	row = 1
	col = 1
	EMSearch "RECIPIENT FILE APPLICATION", row, col
	EMWriteScreen "x", row, col - 3
	transmit
END FUNCTION

FUNCTION navigate_to_MAXIS(maxis_mode)
	the_little_script_that_could = True
	attn
	EMConnect "A"
	IF maxis_mode = "PRODUCTION" THEN
		EMReadScreen prod_running, 7, 6, 15
		IF prod_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect"B"
			attn
			EMReadScreen prod_running, 7, 6, 15
			IF prod_running = "RUNNING" THEN
				x = "B"
			ELSE
				MsgBox "The script is unable to find your session. Please navigate to your MAXIS session and press OK for the script to continue."
				the_little_script_that_could = False
			END IF
		END IF
	ELSEIF maxis_mode = "INQUIRY DB" THEN
		EMReadScreen inq_running, 7, 7, 15
		IF inq_running = "RUNNING" THEN
			x = "A"
		ELSE
			EMConnect "B"
			attn
			EMReadScreen inq_running, 7, 7, 15
			IF inq_running = "RUNNING" THEN
				x = "B"
			ELSE
				MsgBox "The script is unable to find your session. Please navigate to your MAXIS session and press OK for the script to continue."
				the_little_script_that_could = False
			END IF
		END IF
	END IF

	IF the_little_script_that_could = True THEN
		EMConnect (x)
		IF maxis_mode = "PRODUCTION" THEN
			EMWriteScreen "1", 2, 15
			transmit
		ELSEIF maxis_mode = "INQUIRY DB" THEN
			EMWriteScreen "2", 2, 15
			transmit
		END IF
	END IF

END FUNCTION

Function HCRE_panel_bypass()
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 241, 185, "Pull cases into Excel dialog"
  DropListBox 75, 10, 95, 10, "REPT/ACTV"+chr(9)+"REPT/PND2"+chr(9)+"REPT/REVW", screen_to_use
  EditBox 75, 30, 95, 15, x_number
  CheckBox 10, 50, 230, 10, "Check here for all workers in the county (takes several hours to run)", all_workers_check
  CheckBox 30, 75, 55, 10, "ABAWD code", ABAWD_code_check
  CheckBox 30, 90, 75, 10, "All HH membs 19+", all_HH_membs_19_plus_check
  CheckBox 30, 105, 60, 10, "Citizen/Alien ID", imig_ctzn_check
  CheckBox 30, 120, 80, 10, "FS and MFIP Review", FS_MF_review_check
  CheckBox 30, 135, 80, 10, "GA and MSA Review", ga_msa_check
  CheckBox 30, 150, 60, 10, "GRH DOC amt.", grh_doc_amount_check
  CheckBox 30, 165, 50, 10, "HC Review", HC_REVIEW_check
  CheckBox 125, 75, 70, 10, "Homeless Clients", homeless_check
  CheckBox 125, 90, 65, 10, "MAGI Non-MAGI", magi_pct_check
  CheckBox 125, 105, 105, 10, "MAEPD/Part B Reimbursable", maepd_check
  CheckBox 125, 120, 90, 10, "Number of HH membs", number_of_HH_membs_check
  CheckBox 125, 135, 70, 10, "PDED/Rep-Payee", pded_check
  CheckBox 125, 150, 55, 10, "PREG exists", preg_check
  ButtonGroup ButtonPressed
    OkButton 180, 10, 50, 15
    CancelButton 180, 30, 50, 15
  Text 30, 15, 45, 10, "REPT option:"
  Text 15, 35, 60, 10, "Worker to check:"
  GroupBox 10, 65, 220, 115, "Additional items to log"
EndDialog

Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If all_workers_check = 0 and trim(x_number) = "" then err_msg = err_msg & vbcr & "* Select a basket number(s), OR the all workers check box"
        If all_workers_check = 1 and trim(x_number) <> "" then err_msg = err_msg & vbcr & "* Select a basket number(s) OR the all workers check box, not both options."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

IF x_number = "" THEN CALL find_variable("User: ", x_number, 7)

'Adjusting name of script variable for usage stats according to what was done. So, if ACTV was used instead of PND2, it'll indicate that on the script (and thus allow accurate measurement of time savings).
If screen_to_use = "REPT/PND2" then
	name_of_script = "BULK - pull cases into Excel (PND2)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (PND2 all cases)"
ElseIf screen_to_use = "REPT/ACTV" then
	name_of_script = "BULK - pull cases into Excel (ACTV)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (ACTV all cases)"
ElseIf screen_to_use = "REPT/REVW" then
	name_of_script = "BULK - pull cases into Excel (REVW)"
	If all_workers_check = 1 then name_of_script = "BULK - pull cases into Excel (REVW all cases)"
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X Number"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"

'If working off of PND2 it sets the 4th  col as APPL DATE, otherwise it'll be NEXT REVW DATE
If screen_to_use = "REPT/PND2" then
	ObjExcel.Cells(1, 4).Value = "APPL DATE"
ElseIf screen_to_use = "REPT/ACTV" then
	ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"
ElseIf screen_to_use = "REPT/REVW" then
	ObjExcel.Cells(1, 4).Value = "REVW REC'D DATE"
End if

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 4 because cols 1-3 are already used

'----------------------------------------------------------------------------------------------------ABAWD
If ABAWD_code_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "ABAWD CODE"
	ABAWD_code_col = col_to_use
	col_to_use = col_to_use + 1
End if

'----------------------------------------------------------------------------------------------------All HH members 19+
If all_HH_membs_19_plus_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "ALL MEMBS 19+?"
	all_HH_membs_19_plus_col = col_to_use
	col_to_use = col_to_use + 1
End if

'----------------------------------------------------------------------------------------------------Citizenship
IF imig_ctzn_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "Clients with Cit = Y"
	alien_id_col = col_to_use
	col_to_use = col_to_use + 1
	objExcel.Cells(1, col_to_use).Value = "Need Interp?"
	interp_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------FS and MFIP review
IF FS_MF_review_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "SNAP Cases to Review"
	SNAP_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "MFIP Cases to Review"
	MFIP_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------GA and MSA review
IF ga_msa_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "GA Cases to Review"
	ga_revw_col = col_to_use
	col_to_use = col_to_use + 1
	objExcel.Cells(1, col_to_use).Value = "MSA Cases to Review"
	msa_revw_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------GRH DOC amounts
IF grh_doc_amount_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "GRH DOC Amount"
	grh_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------HC cases to review
IF HC_REVIEW_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "HC Cases to Review"
	HC_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------homeless
IF homeless_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "CL reporting homeless?"
	homeless_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------MAGI/Non-MAGI
IF magi_pct_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "MAGI Persons"
	magi_pers_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "Non-MAGI Persons"
	nonmagi_pers_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "# of MAGI"
	magi_count_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "# of Non-MAGI"
	nonmagi_count_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "MAGI Household"
	magi_hh_col = col_to_use
	col_to_use = col_to_use + 1
	ObjExcel.Cells(1, col_to_use).Value = "Mixed Household"
	mixed_hh_col = col_to_use
	col_to_use = col_to_use + 1
	objExcel.Cells(1, col_to_use).Value = "Non-MAGI Household"
	nonmagi_hh_col = col_to_use
	col_to_use = col_to_use + 1
	objExcel.Cells(1, col_to_use).Value = "MAGI Review aligned?"
	reviews_aligned_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------MA-EPD
IF maepd_check = 1 THEN
	objExcel.Cells(1, col_to_use).Value = "MA-EPD & Part B Reimburseable"
	maepd_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------Number of HH members
If number_of_HH_membs_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "NUMBER OF HH MEMBS?"
	number_of_HH_membs_col = col_to_use
	col_to_use = col_to_use + 1
End if

'----------------------------------------------------------------------------------------------------PDED
IF pded_check = 1 THEN
	ObjExcel.Cells(1, col_to_use).Value = "PDED/Rep-Payee"
	pded_col = col_to_use
	col_to_use = col_to_use + 1
END IF

'----------------------------------------------------------------------------------------------------PREG
If preg_check = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "PREG EXISTS?"
	preg_col = col_to_use
	col_to_use = col_to_use + 1
End if

FOR i = 1 TO col_to_use		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

'Setting the variable for what's to come
excel_row = 2

'If all workers are selected, the script will open the worker list stored on the shared drive, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = 1 then
	CALL create_array_of_all_active_x_numbers_in_county(x_array, "27")
Else
	IF len(x_number) > 3 THEN
		x_array = split(x_number, ", ")
	ELSE
		x_array = split(x_number)
	END IF
End if

For each worker in x_array
'Getting to PND2, if PND2 is the selected option
If screen_to_use = "REPT/PND2" then
	Call navigate_to_MAXIS_screen("rept", "pnd2")
	EMWriteScreen worker, 21, 13
	transmit

	'Grabbing each case number on screen
	Do
		row = 7
		Do
			EMReadScreen MAXIS_case_number, 8, row, 5
			If MAXIS_case_number = "        " then
				EMReadScreen additional_app, 14, row, 17
				IF additional_app = "              " THEN
					EXIT DO
				ELSE
					row = row + 1
				END IF
			ELSE
				EMReadScreen client_name, 22, row, 16
				EMReadScreen APPL_date, 8, row, 38
				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
				ObjExcel.Cells(excel_row, 3).Value = client_name
				ObjExcel.Cells(excel_row, 4).Value = replace(APPL_date, " ", "/")
				row = row + 1
			END IF
			Stats_counter = stats_counter + 1
			excel_row = excel_row + 1
		Loop until row = 19
		PF8
		EMReadScreen last_page_check, 21, 24, 2
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if

'Getting to ACTV, if ACTV is the selected option
If screen_to_use = "REPT/ACTV" then
	Call navigate_to_MAXIS_screen("rept", "actv")
	IF worker <> "" THEN
		IF len(worker) = 3 THEN
			EMWriteScreen worker, 21, 17
		ELSE
			EMWriteScreen worker, 21, 13
		END IF
		transmit
	END IF
	EMReadScreen user_id, 7, 21, 71
	EMReadScreen check_worker, 7, 21, 13
	IF user_id = check_worker THEN PF7

	'Grabbing each case number on screen
	Do
		row = 7
		EMReadScreen last_page_check, 21, 24, 2
		Do
			EMReadScreen MAXIS_case_number, 8, row, 12
			If trim(MAXIS_case_number) = "" then exit do
			EMReadScreen client_name, 21, row, 21
			EMReadScreen next_REVW_date, 8, row, 42
			ObjExcel.Cells(excel_row, 1).Value = worker
			ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
			ObjExcel.Cells(excel_row, 3).Value = client_name
			ObjExcel.Cells(excel_row, 4).Value = replace(next_REVW_date, " ", "/")
			row = row + 1
			excel_row = excel_row + 1
			Stats_counter = stats_counter + 1
		Loop until row = 19
		PF8
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if

'Getting to PND2, if PND2 is the selected option
If screen_to_use = "REPT/REVW" then
	Call navigate_to_MAXIS_screen("rept", "revw")
	EMWriteScreen worker, 21, 6
	EMWriteScreen CM_plus_1_mo, 20, 55
	EMWriteScreen CM_plus_1_yr, 20, 58
	transmit

	'Grabbing each case number on screen
	Do
		row = 7
		Do
			EMReadScreen MAXIS_case_number, 7, row, 7
			If trim(MAXIS_case_number) = "" then exit do
			EMReadScreen client_name, 22, row, 16
			EMReadScreen revw_date, 8, row, 62
			ObjExcel.Cells(excel_row, 1).Value = worker
			ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
			ObjExcel.Cells(excel_row, 3).Value = client_name
			If revw_date =  "__ __ __" then
				revw_date = ""
			Else
				revw_date = replace(revw_date, " ", "/")
			End if
			ObjExcel.Cells(excel_row, 4).Value = revw_date
			row = row + 1
			excel_row = excel_row + 1
			Stats_counter = stats_counter + 1
		Loop until row = 19
		PF8
		EMReadScreen last_page_check, 21, 24, 2
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End if
next

'Resetting excel_row variable, now we need to start looking people up
excel_row = 2

Do
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	If MAXIS_case_number = "" then exit do

	'Now pulling PREG info
	If preg_check = 1 then
		call navigate_to_MAXIS_screen("STAT", "PREG")
		EMReadScreen PREG_panel_check, 1, 2, 78
		If PREG_panel_check <> "0" then
			ObjExcel.Cells(excel_row, preg_col).Value = "Y"
		Else
			ObjExcel.Cells(excel_row, preg_col).Value = "N"
		End if
	End if

	'Checking the GRH DOC Amount
	IF grh_doc_amount_check = 1 THEN
		CALL navigate_to_MAXIS_screen("STAT", "FACI")
		grh_doc_amt = ""
		faci_row = 14
		DO
			EMReadScreen open_faci, 10, faci_row, 71
			IF open_faci <> "__ __ ____" THEN
				faci_row = faci_row + 1
				IF faci_row = 19 THEN
					transmit
					faci_row = 14
					EMReadScreen error_message, 15, 24, 2
					error_message = trim(error_message)
					IF error_message <> "" THEN exit do
				END IF
			ELSE
				EMReadScreen start_faci, 10, faci_row, 47
				IF start_faci = "__ __ ____" THEN
					transmit
					faci_row = 14
					EMReadScreen error_message, 15, 24, 2
					error_message = trim(error_message)
					IF error_message <> "" THEN exit do
				ELSE
					EMReadScreen grh_doc_amt, 8, 13, 45
					grh_doc_amt = replace(grh_doc_amt, "_", "")
					ObjExcel.Cells(excel_row, grh_col).Value = grh_doc_amt
					exit do
				END IF
			END IF
		LOOP
	END IF

	'Checking for Citizen = Y and an Alien ID Number
	IF imig_ctzn_check = 1 THEN
		CALL navigate_to_MAXIS_screen("STAT", "MEMB")
		DO
			EMReadScreen needs_trans, 1, 14, 68
			EMReadScreen ref_num, 2, 4, 33
			EMReadScreen alien_id, 10, 15, 68
			alien_id = replace(alien_id, "_", "")

			objExcel.Cells(excel_row, interp_col).Value = objExcel.Cells(excel_row, interp_col).Value & needs_trans & ";"

			EMWriteScreen "MEMI", 20, 71
			EMWriteScreen ref_num, 20, 76
			transmit

			EMReadScreen citizen_yn, 1, 11, 49
			IF citizen_yn = "Y" THEN objExcel.Cells(excel_row, alien_id_col).Value = objExcel.Cells(excel_row, alien_id_col).Value & ref_num & ";"

			EMWriteScreen "MEMB", 20, 71
			EMWriteScreen ref_num, 20, 76
			transmit

			transmit
			EMReadScreen error_message, 20, 24, 2
			error_message = trim(error_message)
		LOOP UNTIL error_message <> ""
		IF InStr(objExcel.Cells(excel_row, interp_col).Value, "Y") <> 0 THEN
			objExcel.Cells(excel_row, interp_col).Value = ""
		ELSE
			objExcel.Cells(excel_row, interp_col).Value = "N"
		END IF
	END IF

	'Now pulling age info
	If all_HH_membs_19_plus_check = 1 then
		call navigate_to_MAXIS_screen("STAT", "MEMB")
		Do
			EMReadScreen MEMB_panel_current, 1, 2, 73
			EMReadScreen MEMB_panel_total, 1, 2, 78
			EMReadScreen MEMB_age, 3, 8, 76
			If MEMB_age = "   " then MEMB_age = "0"
			If cint(MEMB_age) < 19 then has_minor_in_case = True
			transmit
		Loop until MEMB_panel_current = MEMB_panel_total
		If has_minor_in_case <> True then
			ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "Y"
		Else
			ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "N"
		End if
		has_minor_in_case = "" 'clearing variable
	End if

	'Now pulling number of membs info
	If number_of_HH_membs_check = 1 then
		call navigate_to_MAXIS_screen("STAT", "MEMB")
		EMReadScreen MEMB_panel_total, 1, 2, 78
		ObjExcel.Cells(excel_row, number_of_HH_membs_col).Value = cint(MEMB_panel_total)
	End if

	'Now pulling ABAWD info
	If ABAWD_code_check = 1 then
	    'Now pulling ABAWD info
	    ABAWD_status = "" 		'clearing variable
	    eats_group_members = ""		'clearing
	    eats_row = 13			'clearing variable
	    'msgbox MAXIS_case_number
	    call navigate_to_MAXIS_screen("STAT", "PROG")
	    EMReadScreen snap_status, 4, 10, 74
	    IF snap_status = "ACTV" OR snap_status = "PEND" THEN
	    	call HCRE_panel_bypass
	    	call navigate_to_MAXIS_screen("STAT", "EATS")
	    	EMReadScreen all_eat_together, 1, 4, 72
	    	IF all_eat_together = "_" THEN
	    		excel_row = excel_row + 1
	    	ELSEIF all_eat_together = "Y" THEN
	    		eats_row = 5
	    		DO
	    			EMReadScreen eats_person, 2, eats_row, 3
	    			eats_person = trim(eats_person)
	    			IF eats_person <> "" THEN
	    				eats_group_members = eats_group_members & eats_person & " "
	    				eats_row = eats_row + 1
	    			END IF
	    		LOOP UNTIL eats_person = "" or eats_row = 18
	    	ELSEIF all_eat_together = "N" THEN
	    		eats_row = 13
	    		DO
	    			EMReadScreen eats_group, 38, eats_row, 39
	    			find_memb01 = InStr(eats_group, "01")
	    			IF find_memb01 = 0 THEN
	    				eats_row = eats_row + 1
	    			else
	    				exit do
	    			End if
	    		LOOP UNTIL find_memb01 <> 0 OR eats_row = 18
	    		IF eats_row <> 18 THEN
	    			eats_col = 39
	    			DO
	    				EMReadScreen eats_group, 2, eats_row, eats_col
	    				IF eats_group <> "__" THEN
	    					eats_group_members = eats_group_members & eats_group & " "
	    					eats_col = eats_col + 4
	    				END IF
	    			LOOP UNTIL eats_group = "__"
	    		END IF
	    	End if

	    	IF eats_row <> 18 then
	    		eats_group_members = trim(eats_group_members)
	    		eats_group_members = split(eats_group_members)

	    		call navigate_to_MAXIS_screen("STAT", "WREG")

	    		FOR EACH person IN eats_group_members
	    			EMWriteScreen person, 20, 76
	    			transmit
	    			EMReadScreen ABAWD_status_code, 2, 13, 50
	    			ABAWD_status = ABAWD_status & person & ": " & ABAWD_status_code & ","
	    		NEXT

	    		ObjExcel.Cells(excel_row, 5).Value = ABAWD_status
	    		excel_row = excel_row + 1
	    	ELSE
	    		objExcel.Cells(excel_row, 5).Value = "CHECK MANUALLY"
	    		excel_row = excel_row + 1
	    	END IF
	    Else
	    	excel_row = excel_row + 1
		End if
	End if

	IF pded_check = 1 THEN
		total_pded = ""
		pded_hh_array = ""
		call navigate_to_MAXIS_screen("STAT", "PDED")
			pded_row = 5
			DO
				EMReadScreen pded_hh_memb, 2, pded_row, 3
				IF pded_hh_memb = "  " THEN
					EXIT DO
				ELSE
					pded_hh_array = pded_hh_array & pded_hh_memb & " "
					pded_row = pded_row + 1
				END IF
			LOOP UNTIL pded_hh_memb = "  "

			pded_hh_array = trim(pded_hh_array)
			pded_hh_array = split(pded_hh_array)

			FOR EACH hh_memb IN pded_hh_array
				pded_info = ""
				rep_payee_amt = ""
				EMWriteScreen hh_memb, 20, 76
				transmit
					EMReadScreen rep_payee_amt, 8, 15, 70
				rep_payee_amt = replace(rep_payee_amt, "_", "")
				rep_payee_amt = replace(rep_payee_amt, " ", "")
				IF rep_payee_amt <> "" THEN
					pded_info = hh_memb & ": " & rep_payee_amt & "; "
					total_pded = total_pded & pded_info
				END IF
			NEXT

			ObjExcel.Cells(excel_row, PDED_col).Value = total_pded
			'THE FOLLOWING 5 LINES AUTOMATICALLY DELETE ANY BLANK RESULTS
			'IF total_pded = "" THEN
			'	Set objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF

	END IF

	IF magi_pct_check = 1 THEN
		'Finds HC Budget Method
		MAGI_count = 0
		nonMAGI_count = 0
		call navigate_to_MAXIS_screen("ELIG", "HC")
		hhmm_row = 8
		DO
			EMReadScreen hc_ref_num, 2, hhmm_row, 3
			EMReadScreen hc_information_found, 70, hhmm_row, 3
			hc_information_found = trim(hc_information_found)
			EMReadScreen elig_result, 4, hhmm_row, 41
			EMReadScreen elig_status, 6, hhmm_row, 50
			IF hc_information_found <> "" THEN
				IF elig_result = "ELIG" AND elig_status = "ACTIVE" THEN
					EMReadScreen hc_requested, 1, hhmm_row, 28
					IF hc_requested = "S" OR hc_requested = "Q" OR hc_requested = "I" THEN 			'IF the HH MEMB is MSP ONLY then they are automatically Budg Mthd B
						IF hc_ref_num = "  " THEN
							temp_hhmm_row = hhmm_row
							DO
								EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
								IF hc_ref_num = "  " THEN
									temp_hhmm_row = temp_hhmm_row - 1
								ELSE
									EXIT DO
								END IF
							LOOP
						END IF
						IF InStr(objExcel.Cells(excel_row, nonmagi_pers_col).Value, hc_ref_num & ";") = 0 THEN
							ObjExcel.Cells(excel_row, nonmagi_pers_col).Value = ObjExcel.Cells(excel_row, nonmagi_pers_col).Value & hc_ref_num & ";"
							nonMAGI_count = nonMAGI_count + 1
						END IF
						hhmm_row = hhmm_row + 1
					ELSEIF hc_requested = "M" or hc_requested = "E" THEN
						EMWriteScreen "X", hhmm_row, 26
						transmit
						EMReadScreen budg_mthd, 1, 13, 76
						IF budg_mthd = "A" THEN
							IF hc_ref_num = "  " THEN
								temp_hhmm_row = hhmm_row
								DO
									EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
									IF hc_ref_num = "  " THEN
										temp_hhmm_row = temp_hhmm_row - 1
									ELSE
										EXIT DO
									END IF
								LOOP
							END IF
							IF InStr(objExcel.Cells(excel_row, magi_pers_col).Value, hc_ref_num & ";") = 0 THEN
								objExcel.Cells(excel_row, magi_pers_col).Value = ObjExcel.Cells(excel_row, magi_pers_col).Value & hc_ref_num & ";"
								MAGI_count = MAGI_count + 1
							END IF
						ELSE
							IF hc_ref_num = "  " THEN
								temp_hhmm_row = hhmm_row
								DO
									EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
									IF hc_ref_num = "  " THEN
										temp_hhmm_row = temp_hhmm_row - 1
									ELSE
										EXIT DO
									END IF
								LOOP
							END IF
							IF InStr(objExcel.Cells(excel_row, nonmagi_pers_col).Value, hc_ref_num & ";") = 0 THEN
								objExcel.Cells(excel_row, nonmagi_pers_col).Value = ObjExcel.Cells(excel_row, nonmagi_pers_col).Value & hc_ref_num & ";"
								nonMAGI_count = nonMAGI_count + 1
							END IF
						END IF
						PF3
						hhmm_row = hhmm_row + 1
					ELSEIF hc_requested = "N" THEN
						hhmm_row = hhmm_row + 1
					END IF
				ELSE
					hhmm_row = hhmm_row + 1
				END IF
			ELSE
				EXIT DO
			END IF
		LOOP UNTIL hhmm_row = 20 OR hc_ref_num = "  "

		objExcel.Cells(excel_row, magi_count_col).Value = MAGI_count
		objExcel.Cells(excel_row, nonmagi_count_col).Value = nonMAGI_count

		'Checking if client is active on HC then going to find review dates to find MAGI cases that are missing aligned reviews.
		IF MAGI_count <> 0 THEN
			hc_compare_renewal = ""
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			EMReadScreen MAGI_HC_status, 4, 12, 74
			IF MAGI_HC_status = "ACTV" Then
				Call HCRE_panel_bypass
				CALL navigate_to_MAXIS_screen("STAT", "REVW")
				EMwritescreen "X", 5, 71
				Transmit
				'Checking to make sure pop up opened
				DO
					EMReadScreen revw_pop_up_check, 8, 4, 44
					EMWaitReady 1, 1
				LOOP until revw_pop_up_check = "RENEWALS"
				'Reading HC reviews to compare them
				EMReadScreen hc_income_renewal, 8, 8, 27
				EMReadScreen hc_IA_renewal, 8, 8, 71
				EMReadScreen hc_annual_renewal, 8, 9, 27
				IF hc_income_renewal = "__ 01 __" THEN hc_compare_renewal = hc_IA_renewal
				IF hc_IA_renewal = "__ 01 __" THEN hc_compare_renewal = hc_income_renewal
'								If MAXIS_case_number = 302735 THEN msgbox "hc income " & hc_income_renewal & " hc asset " & hc_IA_renewal & " hc annual " & hc_annual_renewal & " compare " & hc_compare_renewal

				IF hc_annual_renewal = hc_compare_renewal THEN
					objExcel.Cells(excel_row, reviews_aligned_col).Value = "Y"
				ELSE
					objExcel.Cells(excel_row, reviews_aligned_col).Value = "Y"
				END IF
			END IF
		END IF

		IF MAGI_count <> 0 AND nonMAGI_count = 0 THEN
			objExcel.Cells(excel_row, magi_hh_col).Value = "Y"
		ELSEIF MAGI_count <> 0 AND nonMAGI_count <> 0 THEN
			objExcel.Cells(excel_row, mixed_hh_col).Value = "Y"
		ELSEIF MAGI_count = 0 AND nonMAGI_count <> 0 THEN
			objExcel.Cells(excel_row, nonmagi_hh_col).Value = "Y"
		'ELSEIF MAGI_count = 0 AND nonMAGI_count = 0 THEN
		'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		'	objRange.Delete
		'	excel_row = excel_row - 1
		END IF
	END IF

	IF FS_MF_review_check = 1 OR HC_REVIEW_check = 1 OR ga_msa_check = 1 THEN
		'resetting variables
		panel_status = ""
		hc_panel_status = ""
		snap_panel_status = ""
		cash_panel_status = ""
		ga_panel_status = ""
		msa_panel_status = ""
		hc_group = ""
		eats_group_members = ""
		mfip_group = ""
		ga_status = ""
		msa_status = ""
		ga_check_array = ""

		CALL navigate_to_MAXIS_screen("STAT", "PROG")
		hc_actv_missing_revw = ""

		EMReadScreen hc_status, 4, 12, 74
		EMReadScreen snap_status, 4, 10, 74
		EMReadScreen CASH_prog_1, 2, 6, 67
		EMReadScreen CASH_prog_2, 2, 7, 67
		EMReadScreen CASH_status_1, 4, 6, 74
		EMReadScreen CASH_status_2, 4, 7, 74

		IF (FS_MF_review_check = 1 AND (snap_status = "ACTV" OR (CASH_prog_1 = "MF" AND CASH_status_1 = "ACTV") OR (CASH_prog_2 = "MF" AND CASH_status_2 = "ACTV"))) OR _
			(HC_REVIEW_check = 1 AND hc_status = "ACTV") OR _
			(ga_msa_check = 1 AND (((CASH_prog_1 = "GA" OR CASH_prog_1 = "MS") AND CASH_status_1 = "ACTV") OR ((CASH_prog_2 = "GA" OR CASH_prog_2 = "MS") AND CASH_status_2 = "ACTV"))) THEN

			Call HCRE_panel_bypass
			CALL navigate_to_MAXIS_screen("STAT", "REVW")
			IF FS_MF_review_check = 1 THEN
				EmReadScreen cash_review_date, 8, 9, 37   'reads cash renewal date
				EMwritescreen "X", 5, 58
				Transmit
				EmReadScreen snap_review_date, 8, 9, 64     'reads snap ER date
				cash_review_date = replace(cash_review_date, " ", "/")
				snap_review_date = replace(snap_review_date, " ", "/")
				PF3
			END IF
			IF HC_REVIEW_check = 1 THEN
				EMWriteScreen "X", 5, 71
				transmit
				EMReadScreen hc_review_date, 8, 9, 27
				IF hc_review_date = "__ 01 __" AND hc_status = "ACTV" THEN   'some cases managed to be actv without a review date entered. Per Pat and PCs this is wrong and we need to catch it.
					hc_review_date = "01 01 00"  'prevents Cdate from erroring out
					hc_actv_missing_revw = true					'Catches method A cases as the logic below would skip them.
				END IF
				hc_review_date = replace(hc_review_date, " ", "/")
				PF3
			END IF
			IF ga_msa_check = 1 THEN
				EMReadScreen cash_review_date, 8, 9, 37
				cash_review_date = replace(cash_review_date, " ", "/")
			END IF
			comparison_date = datepart("M", date) & "/01/" & datepart("yyyy", date)
			past_month = dateadd("M", -1, comparison_date)				'establishes minimum range
			future_month = dateadd("M", 1, comparison_date)				'establishes maximum range

			'calcuate the past current and future renewal months and years
			review_date_1_year_past = dateadd("YYYY", 1, past_month)
			review_date_1_year_current = dateadd("YYYY", 1, comparison_date)
			review_date_1_year_future = dateadd("YYYY", 1, future_month)
			review_date_2_year_past = dateadd("YYYY", 2, past_month)
			review_date_2_year_current = dateadd("YYYY", 2, comparison_date)
			review_date_2_year_future = dateadd("YYYY", 2, future_month)

			IF HC_REVIEW_check = 1 AND hc_status = "ACTV" THEN
				IF (cdate(hc_review_date) = cdate(review_date_1_year_past) OR _
					cdate(hc_review_date) = cdate(review_date_1_year_current) OR _
					cdate(hc_review_date) = cdate(review_date_1_year_future) OR _
					cdate(hc_review_date) = cdate(review_date_2_year_past) OR _
					cdate(hc_review_date) = cdate(review_date_2_year_current) OR _
					cdate(hc_review_date) = cdate(review_date_2_year_future)) THEN

					' >>>>> Creating an array of all active HC members.
					' >>>>> Per conversation with Pat B on 7/21/15, we are going to look
					' >>>>>>>> only for HH members with a Budget Method = B, L, or S that have
					' >>>>>>>> the following panels...
					' >>>>>>>>		JOBS	UNEA	BUSI	RBIC
					' >>>>>>>>		ACCT	SECU	OTHR	CASH
					' >>>>>>>>		CARS	REST	SPON(?)

					CALL navigate_to_MAXIS_screen("ELIG", "HC")
					HHMM_row = 8
					hc_group = ""
					DO
						' >>>>> Grabbing the HH reference number and the HC program.
						EMReadScreen hhmm_ref_num, 2, HHMM_row, 3
						hhmm_ref_num = replace(hhmm_ref_num, " ", "")
						EMReadScreen hhmm_program, 10, HHMM_row, 28
						hhmm_program = replace(hhmm_program, " ", "")
						IF hhmm_ref_num <> "" AND hhmm_program <> "" THEN
							IF hhmm_program = "SLMB" OR hhmm_program = "QI1" OR hhmm_program = "QMB" THEN
								IF InStr(hc_group, hhmm_ref_num) = 0 THEN hc_group = hc_group & hhmm_ref_num & "~~~"
							ELSEIF (hhmm_program <> "QMB" OR hhmm_program <> "SLMB" OR hhmm_program <> "QI1") AND (hhmm_program <> "NO REQUEST" AND hhmm_program <> "NO VERSION") THEN
								EMWriteScreen "X", HHMM_row, 26
								transmit
								EMReadScreen hc_budget_mthd, 1, 13, 76
								IF hc_budget_mthd = "B" OR hc_budget_mthd = "L" OR hc_budget_mthd = "S" THEN
									IF InStr(hc_group, hhmm_ref_num) = 0 THEN hc_group = hc_group & hhmm_ref_num & "~~~"
								END IF
								PF3
							END IF
						ELSEIF hhmm_ref_num = "" AND hhmm_program <> "" THEN
							IF hhmm_program = "SLMB" OR hhmm_program = "QI1" OR hhmm_program = "QMB" THEN
								hhmm_row_back_up = 1
								DO
									EMReadScreen hh_member, 2, HHMM_row - hhmm_row_back_up, 3
									IF hh_member = "  " THEN hhmm_row_back_up = hhmm_row_back_up + 1
								LOOP UNTIL hh_member <> "  " OR hhmm_row_back_up = 3
								IF InStr(hc_group, hh_member) = 0 AND hh_member <> "  " THEN hc_group = hc_group & hh_member & "~~~"
							END IF
						ELSEIF hhmm_ref_num = "" AND hhmm_program = "" THEN
							PF3
						END IF
						HHMM_row = HHMM_row + 1
					LOOP UNTIL hhmm_ref_num = "" AND hhmm_program = ""
					back_to_SELF
					hc_group = trim(hc_group)
					hc_group = split(hc_group, "~~~")

					CALL check_panels_for_hc_function(hc_group, panel_status)
					hc_panel_status = panel_status
					IF hc_panel_status <> "" THEN objExcel.Cells(excel_row, HC_col).Value = "REVIEW HC"
				END IF
				IF hc_actv_missing_revw = true THEN objExcel.Cells(excel_row, HC_col).Value = "REVIEW HC (review dates missing)"   'marks the cases with HC actv and no renewal dates as needing review.
			END IF

			IF FS_MF_review_check = 1 AND snap_status = "ACTV" THEN
				IF (cdate(snap_review_date) = cdate(review_date_1_year_past) OR _
					cdate(snap_review_date) = cdate(review_date_1_year_current) OR _
					cdate(snap_review_date) = cdate(review_date_1_year_future) OR _
					cdate(snap_review_date) = cdate(review_date_2_year_past) OR _
					cdate(snap_review_date) = cdate(review_date_2_year_current) OR _
					cdate(snap_review_date) = cdate(review_date_2_year_future)) THEN

					call navigate_to_MAXIS_screen("STAT", "EATS")
					EMReadScreen all_eat_together, 1, 4, 72
					IF all_eat_together = "_" THEN
						eats_group_members = "01" & " "
					ELSEIF all_eat_together = "Y" THEN
						eats_row = 5
						DO
							EMReadScreen eats_person, 2, eats_row, 3
							eats_person = replace(eats_person, " ", "")
							'IF instr(eats_person, "?") <> 0 THEN
								IF eats_person <> "" THEN
									eats_group_members = eats_group_members & eats_person & " "
									eats_row = eats_row + 1
								END IF
							'END IF
						LOOP UNTIL eats_person = ""
					ELSEIF all_eat_together = "N" THEN
						eats_row = 13
						DO
							EMReadScreen eats_group, 38, eats_row, 39
							find_memb01 = InStr(eats_group, "01")
							IF find_memb01 = 0 THEN eats_row = eats_row + 1
						LOOP UNTIL find_memb01 <> 0
						eats_col = 39
						DO
							EMReadScreen eats_group, 2, eats_row, eats_col
							IF eats_group <> "__" THEN
								eats_group_members = eats_group_members & eats_group & " "
								eats_col = eats_col + 4
							END IF
						LOOP UNTIL eats_group = "__"
					END IF

					eats_group_members = trim(eats_group_members)
					eats_group_members = split(eats_group_members)

					CALL check_panels_function(eats_group_members, panel_status)
					snap_panel_status = panel_status
					IF snap_panel_status <> "" THEN ObjExcel.Cells(excel_row, snap_col).Value = "Review SNAP"
				END IF
			END IF
			IF FS_MF_review_check = 1 AND ((CASH_prog_1 = "MF" AND CASH_status_1 = "ACTV") OR (CASH_prog_2 = "MF" AND CASH_status_2 = "ACTV")) THEN
				IF (cdate(cash_review_date) = cdate(review_date_1_year_past) OR _
					cdate(cash_review_date) = cdate(review_date_1_year_current) OR _
					cdate(cash_review_date) = cdate(review_date_1_year_future) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_past) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_current) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_future)) THEN

					panel_status = ""

					CALL navigate_to_MAXIS_screen("ELIG", "MFIP")
					mfpr_row = 7
					DO
						IF mfpr_row = 18 THEN
							PF8
							EMReadScreen no_more_members, 15, 24, 5
							mfpr_row = 7
						END IF
						EMReadScreen is_counted, 7, mfpr_row, 41
						is_counted = replace(is_counted, " ", "")
						IF is_counted = "COUNTED" THEN
							EMReadScreen ref_num, 2, mfpr_row, 6
							mfip_group = mfip_group & ref_num & " "
						END IF
						mfpr_row = mfpr_row + 1
					LOOP UNTIL is_counted = "" OR no_more_members = "NO MORE MEMBERS"
					mfip_group = trim(mfip_group)
					mfip_group = split(mfip_group)

					CALL check_panels_function(mfip_group, panel_status)
					cash_panel_status = panel_status
					IF cash_panel_status <> "" THEN ObjExcel.Cells(excel_row, mfip_col).Value = "Review MFIP"
				END IF
			END IF
			IF ga_msa_check = 1 AND (((CASH_prog_1 = "GA" OR CASH_prog_1 = "MS") AND CASH_status_1 = "ACTV") OR ((CASH_prog_2 = "GA" OR CASH_prog_2 = "MS") AND CASH_status_2 = "ACTV")) THEN
				IF (cdate(cash_review_date) = cdate(review_date_1_year_past) OR _
					cdate(cash_review_date) = cdate(review_date_1_year_current) OR _
					cdate(cash_review_date) = cdate(review_date_1_year_future) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_past) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_current) OR _
					cdate(cash_review_date) = cdate(review_date_2_year_future)) THEN

					panel_status = ""

					'Finding GA folks
					IF (CASH_prog_1 = "GA" AND CASH_status_1 = "ACTV") OR (CASH_prog_2 = "GA" AND CASH_status_2 = "ACTV") THEN
						objExcel.Cells(excel_row, ga_revw_col).Value = "REVW GA"
						ga_panel_status = "REVW"
						CALL navigate_to_MAXIS_screen("ELIG", "GA")
						gapr_row = 8
						DO
							EMReadScreen cash_elig, 4, gapr_row, 57
							IF cash_elig = "ELIG" THEN
								EMReadScreen ref_num, 2, gapr_row, 9
								ga_check_array = ga_check_array & ref_num & ","
							END IF
							gapr_row = gapr_row + 1
							IF gapr_row = 18 THEN
								PF8
								gapr_row = 8
							END IF
						LOOP UNTIL cash_elig = "    "

						ga_check_array = ga_check_array & "END"
						ga_check_array = replace(ga_check_array, ",END", "")
						ga_check_array = split(ga_check_array, ",")

						CALL check_panels_for_income(ga_check_array, panel_status)
						IF panel_status <> "" THEN objExcel.Cells(excel_row, ga_revw_col).Interior.ColorIndex = 6
					END IF

					'Finding the MSA folks
					IF (CASH_prog_1 = "MS" AND CASH_status_1 = "ACTV") OR (CASH_prog_2 = "MS" AND CASH_status_2 = "ACTV") THEN
						objExcel.Cells(excel_row, msa_revw_col).Value = "REVW MSA"
						msa_panel_status = "REVW"
					END IF
				END IF
			END IF
		END IF
		'cleaning up blank rows
		'If revewing FS/MFIP AND HC AND GA/MSA...
		IF FS_MF_review_check = 1 AND HC_REVIEW_check = 1 AND ga_msa_check = 1 THEN
			'IF hc_panel_status = "" AND snap_panel_status = "" AND cash_panel_status = "" and hc_actv_missing_revw = "" AND ga_panel_status = "" AND msa_panel_status = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		'If reviewing FS/MFIP and HC...
		ELSEIF FS_MF_review_check = 1 AND HC_REVIEW_check = 1 AND ga_msa_check = 0 THEN
			'IF hc_panel_status = "" AND snap_panel_status = "" AND cash_panel_status = "" and hc_actv_missing_revw = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		'If reviewing FS/MFIP and GA/MSA...
		ELSEIF FS_MF_review_check = 1 AND HC_REVIEW_check = 0 AND ga_msa_check = 1 THEN
			'IF snap_panel_status = "" AND cash_panel_status = "" AND ga_panel_status = "" AND msa_panel_status = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		'If reviewing HC and GA/MSA...
		ELSEIF FS_MF_review_check = 0 AND HC_REVIEW_check = 1 AND ga_msa_check = 1 THEN
			'IF hc_panel_status = "" AND hc_actv_missing_revw = "" AND ga_panel_status = "" AND msa_panel_status = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		'If reviewing GA/MSA only...
		ELSEIF FS_MF_review_check = 0 AND HC_REVIEW_check = 0 AND ga_msa_check = 1 THEN
			'IF ga_panel_status = "" AND msa_panel_status = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		'If reviewing HC only...
		ELSEIF FS_MF_review_check = 0 AND HC_REVIEW_check = 1 AND ga_msa_check = 0 THEN
			'IF hc_panel_status = "" AND hc_actv_missing_revw = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		'If reviewing FS/MFIP only...
		ELSEIF FS_MF_review_check = 1 AND HC_REVIEW_check = 0 AND ga_msa_check = 0 THEN
			'IF snap_panel_status = "" AND cash_panel_status = "" THEN
			'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
			'	objRange.Delete
			'	excel_row = excel_row - 1
			'END IF
		END IF
	END IF

	IF homeless_check = 1 THEN
		CALL navigate_to_MAXIS_screen("STAT", "ADDR")
		EMReadScreen addr_line, 16, 6, 43
		EMReadScreen homeless_yn, 1, 10, 43
		IF homeless_yn = "Y" OR addr_line = "GENERAL DELIVERY" THEN
			objExcel.Cells(excel_row, homeless_col).Value = "HOMELESS"
		'ELSEIF homeless_yn <> "Y" AND addr_line <> "GENERAL DELIVERY" THEN
		'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		'	objRange.Delete
		'	excel_row = excel_row - 1
		END IF
	END IF

	IF MAEPD_check = 1 THEN
		back_to_SELF
		CALL find_variable("Environment: ", production_or_inquiry, 10)
		CALL navigate_to_MAXIS_screen("ELIG", "HC")
		hhmm_row = 8
		DO
			EMReadScreen hc_type, 2, hhmm_row, 28
			IF hc_type = "MA" THEN
				EMWriteScreen "X", hhmm_row, 26
				transmit
				EMReadScreen elig_type, 2, 12, 72
				IF elig_type = "DP" THEN
					EMWriteScreen "X", 9, 76
					transmit
					EMReadScreen pct_fpg, 4, 18, 38
					pct_fpg = trim(pct_fpg)
					pct_fpg = pct_fpg * 1
					IF pct_fpg < 201 THEN
						PF3
						PF3
						EMReadScreen hh_memb_num, 2, hhmm_row, 3
						CALL navigate_to_MAXIS_screen("STAT", "MEMB")
						EMWriteScreen hh_memb_num, 20, 76
						transmit
						EMReadScreen cl_pmi, 8, 4, 46
						cl_pmi = replace(cl_pmi, " ", "")
						DO
							IF len(cl_pmi) <> 8 THEN cl_pmi = "0" & cl_pmi
						LOOP UNTIL len(cl_pmi) = 8
						navigate_to_MMIS
						DO
							EMReadScreen RKEY, 4, 1, 52
							IF RKEY <> "RKEY" THEN EMWaitReady 0, 0
						LOOP UNTIL RKEY = "RKEY"
						EMWriteScreen "I", 2, 19
						EMWriteScreen cl_pmi, 4, 19
						transmit
						EMWriteScreen "RELG", 1, 8
						transmit

						'Reading RELG to determine if the CL is active on MA-EPD
						EMReadScreen prog01_type, 8, 6, 13
							EMReadScreen elig01_type, 2, 6, 33
							EMReadScreen elig01_end, 8, 7, 36
						EMReadScreen prog02_type, 8, 10, 13
							EMReadScreen elig02_type, 2, 10, 33
							EMReadScreen elig02_end, 8, 11, 36
						EMReadScreen prog03_type, 8, 14, 13
							EMReadScreen elig03_type, 2, 14, 33
							EMReadScreen elig03_end, 8, 15, 36
						EMReadScreen prog04_type, 8, 18, 13
							EMReadScreen elig04_type, 2, 18, 33
							EMReadScreen elig04_end, 8, 19, 36

						IF ((prog01_type = "MEDICAID" AND elig01_type = "DP" AND elig01_end = "99/99/99") OR _
							(prog02_type = "MEDICAID" AND elig02_type = "DP" AND elig02_end = "99/99/99") OR _
							(prog03_type = "MEDICAID" AND elig03_type = "DP" AND elig03_end = "99/99/99") OR _
							(prog04_type = "MEDICAID" AND elig04_type = "DP" AND elig04_end = "99/99/99")) THEN

							EMWriteScreen "RMCR", 1, 8
							transmit

							'-----CHECKING FOR ON-GOING MEDICARE PART B-----
							EMReadScreen part_b_begin01, 8, 13, 4
								part_b_begin01 = trim(part_b_begin01)
							EMReadScreen part_b_end01, 8, 13, 15
							EMReadScreen part_b_begin02, 8, 14, 4
								part_b_begin02 = trim(part_b_begin02)
							EMReadScreen part_b_end02, 8, 14, 15

							IF (part_b_begin01 <> "" AND part_b_end01 = "99/99/99") THEN
								EMWriteScreen "RBYB", 1, 8
								transmit

								EMReadScreen accrete_date, 8, 5, 66
								EMReadScreen delete_date, 8, 6, 65
								accrete_date = replace(accrete_date, " ", "")

								IF ((accrete_date = "") OR (accrete_date <> "" AND delete_date <> "99/99/99")) THEN
									objExcel.Cells(excel_row, maepd_col).Value = objExcel.Cells(excel_row, maepd_col).Value & ("MEMB " & hh_memb_num & " ELIG FOR REIMBURSEMENT, ")
								END IF
								PF3
							END IF
						ELSE
							PF3
						END IF
						CALL navigate_to_MAXIS(production_or_inquiry)
						hhmm_row = hhmm_row + 1
						CALL navigate_to_MAXIS_screen("ELIG", "HC")
					ELSE
						DO
							EMReadScreen at_hhmm, 4, 3, 51
							IF at_hhmm <> "HHMM" THEN PF3
						LOOP UNTIL at_hhmm = "HHMM"
						hhmm_row = hhmm_row + 1
					END IF
				ELSE
					PF3
					hhmm_row = hhmm_row + 1
				END IF
			ELSE
				hhmm_row = hhmm_row + 1
			END IF
			IF hhmm_row = 20 THEN
				PF8
				EMReadScreen this_is_the_last_page, 21, 24, 2
			END IF
		LOOP UNTIL hc_type = "  " OR this_is_the_last_page = "THIS IS THE LAST PAGE"
		'Deleting the blank results to clean up the spreadsheet
		'IF objExcel.Cells(excel_row, maepd_col).Value = "" THEN
		'	SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		'	objRange.Delete
		'	excel_row = excel_row - 1
		'END IF
	END IF
	excel_row = excel_row + 1
Loop until MAXIS_case_number = ""

IF ga_msa_check = 1 THEN
	objExcel.Cells(1, col_to_use + 1).Value = "GA cases highlighted YELLOW indicate GA case with active income panel."
	objExcel.Cells(1, col_to_use + 1).Interior.ColorIndex = 6
	objExcel.Columns(col_to_use + 1).AutoFit()
END IF

stats_counter = stats_counter - 1 'started at 1, minus 1 to ensure counts are correct
col_to_use = col_to_use + 2

'Query date/time/runtime info
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(3, col_to_use - 1).Value = "Case count:"	'Goes back one, as this is on the next row
objExcel.Cells(3, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time
ObjExcel.Cells(3, col_to_use).Value = STATS_counter

FOR i = 1 TO col_to_use
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("DONE!!")

'----------------------------------------------------------------------------------------------------Excess code, not being used right now
'IF banked_month_check = 1 THEN
'	objExcel.Cells(1, col_to_use).Value = "Banked Months on MISC?"
'	banked_col = col_to_use
'	col_to_use = col_to_use + 1
'END IF

'IF banked_month_check = 1 THEN
'	CALL navigate_to_MAXIS_screen("STAT", "MISC")
'	EMReadScreen num_of_MISC, 1, 2, 78
'	IF num_of_MISC = "1" THEN
'		MISC_row = 6
'		DO
'			EMReadScreen MISC_description, 25, MISC_row, 30
'			EMReadScreen MISC_item, 10, MISC_row, 66
'			IF InStr(UCase(MISC_description), "BANKED") <> 0 THEN
'				MISC_description = replace(MISC_description, "_", "")
'				MISC_item = replace(MISC_item, "_", "")
'				objExcel.Cells(excel_row, banked_col).Value = objExcel.Cells(excel_row, banked_col).Value & MISC_description & "," & MISC_item & "; "
'			END IF
'			MISC_row = MISC_row + 1
'		LOOP UNTIL MISC_row = 17
'	END IF
'	IF objExcel.Cells(excel_row, banked_col).Value = "" THEN
'		SET objBankedRange = objExcel.Cells(excel_row, 1).EntireRow
'		objBankedRange.Delete
'		excel_row = excel_row - 1
'	END IF
'END IF

'IF imig_ctzn_check = 1 THEN
'	excel_row = 2
'	DO
'		back_to_SELF
'		MAXIS_case_number = objExcel.Cells(excel_row, 2).Value
'		IF objExcel.Cells(excel_row, interp_col).Value = "N" AND objExcel.Cells(excel_row, alien_id_col).Value <> "" THEN
'			membs_group = objExcel.Cells(excel_row, alien_id_col).Value
'			membs_group = split(membs_group, ";")
'			actv_group = ""
'
'			FOR EACH person IN membs_group
'				IF person <> "" THEN
'					CALL navigate_to_MAXIS_screen("CASE", "PERS")
'					pers_row = 9
'					DO
'						EMReadScreen pers_pers, 2, pers_row, 3
'						IF pers_pers = person THEN
'							EMReadScreen active_progs, 34, pers_row, 46
'							IF InStr(active_progs, "A") <> 0 THEN actv_group = actv_group & person & ";"
'						ELSE
'							pers_row = pers_row + 1
'							IF pers_row = 18 THEN
'								PF8
'								pers_row = 9
'								EMReadScreen last_page, 21, 24, 2
'								IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
'							END IF
'						END IF
'					LOOP UNTIL last_page = "THIS IS THE LAST PAGE" OR pers_pers = person
'				END IF
'			NEXT
'			objExcel.Cells(excel_row, alien_id_col).Value = actv_group
'		END IF
'		excel_row = excel_row + 1
'	LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""
'END IF
