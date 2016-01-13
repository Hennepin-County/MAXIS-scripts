'Built by Robert Kalb and Charles Potter of Anoka County

'Gathering stats
name_of_script = "ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 98                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 166, 70, "Case number dialog"
  EditBox 65, 5, 70, 15, MAXIS_case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 130, 25, 30, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 50, 50, 15
    CancelButton 95, 50, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 100, 30, 25, 10, "Year:"
EndDialog

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
footer_month = datepart("m", date)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = Cstr(right(DatePart("YYYY", date), 2))
cstr(footer_month)

EMConnect ""
CALL check_for_MAXIS(False)

CALL MAXIS_case_number_finder(MAXIS_case_number)
call find_variable("Month: ", footer_month, 2)
If row <> 0 then 
  footer_month = footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = footer_year
End if

cstr(footer_month)

DO
	err_msg = ""
	DIALOG case_number_dialog
		cancel_confirmation
		IF MAXIS_case_number = "" THEN err_msg = err_msg & vbCr & "* Please enter a case number."
		IF footer_month = "" THEN err_msg = err_msg & vbCr & "* Please enter a benefit month."
		IF footer_year = "" THEN err_msg = err_msg & vbCr & "* Please enter a benefit year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""
case_number = MAXIS_case_number
CALL check_for_MAXIS(False)

back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
'>>>>>Checking for privileged<<<<<
row = 1 
col = 1 
EMSearch "PRIVILEGED", row, col
IF row <> 0 THEN script_end_procedure("This case appears to be privileged. The script cannot access it.")


DO
	CALL HH_member_custom_dialog(HH_member_array)
	IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
LOOP UNTIL uBound(HH_member_array) <> -1

'Building a placeholder array for EATS group comparison
placeholder_HH_array = ""
person_count = 0
FOR EACH person IN HH_member_array
	placeholder_HH_array = placeholder_HH_array & person & ","
NEXT
	
CALL check_for_MAXIS(False)

closing_message = ""

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
FOR EACH person IN HH_member_array
	IF person <> "" THEN 
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen cl_age, 2, 8, 76
		cl_age = cl_age * 1
		IF cl_age < 18 OR cl_age >= 50 THEN closing_message = closing_message & vbCr & "* Household Member " & person & " appears to have exemption. Age = " & cl_age & "."
	END IF
NEXT

CALL navigate_to_MAXIS_screen("STAT", "DISA")
FOR EACH person IN HH_member_array
	disa_status = false
	IF person <> "" THEN 
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_DISA, 1, 2, 78
		IF num_of_DISA <> "0" THEN 
			EMReadScreen disa_end_dt, 10, 6, 69
			disa_end_dt = replace(disa_end_dt, " ", "/")
			EMReadScreen cert_end_dt, 10, 7, 69
			cert_end_dt = replace(cert_end_dt, " ", "/")
			IF IsDate(disa_end_dt) = True THEN 
				IF DateDiff("D", date, disa_end_dt) > 0 THEN 
					closing_message = closing_message & vbCr & "* Household member " & person & " appears to have disability exemption. DISA end date = " & disa_end_dt & "."
					disa_status = True
				END IF
			ELSE
				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN 
					closing_message = closing_message & vbCr & "* Household member " & person & " appears to have disability exemption. DISA has no end date."
					disa_status = True
				END IF
			END IF
			IF IsDate(cert_end_dt) = True AND disa_status = False THEN 
				IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have disability exemption. DISA Certification end date = " & cert_end_dt & "."
			ELSE
				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN 
					EMReadScreen cert_begin_dt, 8, 7, 47
					IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have disability exemption. DISA certification has no end date."
				END IF
			END IF
		END IF
	END IF
NEXT
		
		
'>>>>>>>>>>>> EATS GROUP
FOR EACH person IN HH_member_array
	CALL navigate_to_MAXIS_screen("STAT", "EATS")
	eats_group_members = ""
	memb_found = True
	EMReadScreen all_eat_together, 1, 4, 72
	IF all_eat_together = "_" THEN
		eats_group_members = "01" & ","
	ELSEIF all_eat_together = "Y" THEN 
		eats_row = 5
		DO
			EMReadScreen eats_person, 2, eats_row, 3
			eats_person = replace(eats_person, " ", "")
			IF eats_person <> "" THEN 
				eats_group_members = eats_group_members & eats_person & ","
				eats_row = eats_row + 1
			END IF
		LOOP UNTIL eats_person = ""
	ELSEIF all_eat_together = "N" THEN
		eats_row = 13
		DO
			EMReadScreen eats_group, 38, eats_row, 39
			find_memb01 = InStr(eats_group, person)
			IF find_memb01 = 0 THEN 
				eats_row = eats_row + 1
				IF eats_row = 18 THEN 
					memb_found = False
					EXIT DO
				END IF
			END IF
		LOOP UNTIL find_memb01 <> 0
		eats_col = 39
		DO
			EMReadScreen eats_group, 2, eats_row, eats_col
			IF eats_group <> "__" THEN 
				eats_group_members = eats_group_members & eats_group & ","	
				eats_col = eats_col + 4
			END IF
		LOOP UNTIL eats_group = "__"
	END IF
	
	IF memb_found = True THEN 
		IF placeholder_HH_array <> eats_group_members THEN script_end_procedure("You are asking the script to verify ABAWD and SNAP E&T exemptions for a household that does not match the EATS group. The script cannot support this request. It will now end." & vbCr & vbCr & "Please re-run the script selecting only the individuals in the EATS group.")
		eats_group_members = trim(eats_group_members)
		eats_group_members = split(eats_group_members, ",")
	
		IF all_eat_together <> "_" THEN 
			CALL write_value_and_transmit("MEMB", 20, 71)
			FOR EACH eats_pers IN eats_group_members
				IF eats_pers <> "" AND person <> eats_pers THEN
					CALL write_value_and_transmit(eats_pers, 20, 76)
					EMReadScreen cl_age, 2, 8, 76
					IF cl_age <> "  " THEN 
						cl_age = cl_age * 1
						IF cl_age =< 17 THEN 
							closing_message = closing_message & vbCr & "* Household member " & person & " may have exemption for minor child caretaker. Household member " & eats_pers & " is minor. Please review for accuracy."
						END IF
					END IF
				END IF
			NEXT
		END IF
		
		CALL write_value_and_transmit("DISA", 20, 71)
		FOR EACH disa_pers IN eats_group_members
			disa_status = false
			IF disa_pers <> "" AND disa_pers <> person THEN 
				CALL write_value_and_transmit(disa_pers, 20, 76)
				EMReadScreen num_of_DISA, 1, 2, 78
				IF num_of_DISA <> "0" THEN 
					EMReadScreen disa_end_dt, 10, 6, 69
					disa_end_dt = replace(disa_end_dt, " ", "/")
					EMReadScreen cert_end_dt, 10, 7, 69
					cert_end_dt = replace(cert_end_dt, " ", "/")
					IF IsDate(disa_end_dt) = True THEN 
						IF DateDiff("D", date, disa_end_dt) > 0 THEN 
							closing_message = closing_message & vbCr & "* Household member " & person & " appears to have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
							disa_status = TRUE
						END IF
					ELSEIF IsDate(disa_end_dt) = False THEN 
						IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN 
							closing_message = closing_message & vbCr & "* Household member " & person & " appears to have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
							disa_status = true
						END IF
					END IF
					IF IsDate(cert_end_dt) = True AND disa_status = False THEN 
						IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have exemption for disabled household member. Member " & disa_pers & " DISA certification end date = " & cert_end_dt & "."
					ELSE
						IF (cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999") THEN 
							EMReadScreen cert_begin_dt, 8, 7, 47
							IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have exemption for disabled household member. Member " & disa_pers & " DISA certification has no end date."
						END IF
					END IF
				END IF				
			END IF
		NEXT
	END IF
NEXT

'>>>>>>>>>>>>>>EARNED INCOME			
FOR EACH person IN HH_member_array			
	IF person <> "" THEN 
		prosp_inc = 0
		prosp_hrs = 0
				
		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_JOBS, 1, 2, 78
		IF num_of_JOBS <> "0" THEN 
			DO
				EMReadScreen jobs_end_dt, 8, 9, 49
				EMReadScreen cont_end_dt, 8, 9, 73
				IF jobs_end_dt = "__ __ __" THEN 
					CALL write_value_and_transmit("X", 19, 38)
					EMReadScreen prosp_monthly, 8, 18, 56
					prosp_monthly = trim(prosp_monthly)
					IF prosp_monthly = "" THEN prosp_monthly = 0
					prosp_inc = prosp_inc + prosp_monthly
					EMReadScreen prosp_hrs, 8, 16, 50
					IF prosp_hrs = "        " THEN prosp_hrs = 0
					pp_hrs = pp_hrs * 1
					EMReadScreen pay_freq, 1, 5, 64
					IF pay_freq = "1" THEN 
						prosp_hrs = prosp_hrs + pp_hrs
					ELSEIF pay_freq = "2" THEN 
						prosp_hrs = prosp_hrs + (2 * pp_hrs)
					ELSEIF pay_freq = "3" THEN 
						prosp_hrs = prosp_hrs + (2.15 * pp_hrs)			
					ELSEIF pay_freq = "4" THEN 
						prosp_hrs = prosh_hrs + (4.3 * pp_hrs)
					END IF
				ELSE
					jobs_end_dt = replace(jobs_end_dt, " ", "/")
					IF DateDiff("D", date, jobs_end_dt) > 0 THEN 
						'Going into the PIC for a job with an end date in the future
						CALL write_value_and_transmit("X", 19, 38)
						EMReadScreen prosp_monthly, 8, 18, 56
						prosp_monthly = trim(prosp_monthly)
						IF prosp_monthly = "" THEN prosp_monthly = 0
						prosp_inc = prosp_inc + prosp_monthly
						EMReadScreen prosp_hrs, 8, 16, 50
						IF prosp_hrs = "        " THEN prosp_hrs = 0
						pp_hrs = pp_hrs * 1
						EMReadScreen pay_freq, 1, 5, 64
						IF pay_freq = "1" THEN 
							prosp_hrs = prosp_hrs + pp_hrs
						ELSEIF pay_freq = "2" THEN 
							prosp_hrs = prosp_hrs + (2 * pp_hrs)
						ELSEIF pay_freq = "3" THEN 
							prosp_hrs = prosp_hrs + (2.15 * pp_hrs)			
						ELSEIF pay_freq = "4" THEN 
							prosp_hrs = prosp_hrs + (4.3 * pp_hrs)
						END IF
					END IF
				END IF
				transmit
				transmit
				EMReadScreen enter_a_valid_command, 13, 24, 2
			LOOP UNTIL enter_a_valid_command = "ENTER A VALID"
		END IF	
		
		EMWriteScreen "BUSI", 20, 71
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_BUSI, 1, 2, 78
		IF num_of_BUSI <> "0" THEN 
			DO
				EMReadScreen busi_end_dt, 8, 5, 72
				busi_end_dt = replace(busi_end_dt, " ", "/")
				IF IsDate(busi_end_dt) = True THEN 
					IF DateDiff("D", date, busi_end_dt) > 0 THEN 
						EMReadScreen busi_inc, 8, 10, 69
						busi_inc = trim(busi_inc)
						EMReadScreen busi_hrs, 3, 13, 74
						busi_hrs = trim(busi_hrs)
						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
						prosp_inc = prosp_inc + busi_inc
						prosp_hrs = prosp_hrs + busi_hrs
					END IF
				ELSE
					IF busi_end_dt = "__/__/__" THEN 
						EMReadScreen busi_inc, 8, 10, 69
						busi_inc = trim(busi_inc)
						EMReadScreen busi_hrs, 3, 13, 74
						busi_hrs = trim(busi_hrs)
						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
						prosp_inc = prosp_inc + busi_inc
						prosp_hrs = prosp_hrs + busi_hrs
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		END IF
		
		EMWriteScreen "RBIC", 20, 71
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_RBIC, 1, 2, 78
		IF num_of_RBIC <> "0" THEN closing_message = closing_message & vbCr & "* Household member " & person & " has RBIC panel. Please review for ABAWD and/or SNAP E&T exemption."
	
		IF prosp_inc >= 935.25 OR prosp_hrs >= 129 THEN 
			closing_message = closing_message & vbCr & "* Household member " & person & " appears to be earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions."
		ELSEIF prosp_hrs >= 80 AND prosp_hrs < 129 THEN 
			closing_message = closing_message & vbCr & "* Household member " & person & " appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption."
		END IF
	END IF
NEXT
		
'>>>>>>>>>>>>UNEA		
CALL navigate_to_MAXIS_screen("STAT", "UNEA")			
FOR EACH person IN HH_member_array
	IF person <> "" THEN 
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_UNEA, 1, 2, 78
		IF num_of_UNEA <> "0" THEN 
			DO
				EMReadScreen unea_type, 2, 5, 37
				EMReadScreen unea_end_dt, 8, 7, 68
				unea_end_dt = replace(unea_end_dt, " ", "/")
				IF IsDate(unea_end_dt) = True THEN 
					IF DateDiff("D", date, unea_end_dt) > 0 THEN
						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
					END IF
				ELSE
					IF unea_end_dt = "__/__/__" THEN 
						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		END IF
	END IF
NEXT
			
'>>>>>>>>>PBEN
CALL navigate_to_MAXIS_screen("STAT", "PBEN")
FOR EACH person IN HH_member_array
	IF person <> "" THEN 
		EMWriteScreen "PBEN", 20, 71
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_PBEN, 1, 2, 78
		IF num_of_PBEN <> "0" THEN 
			pben_row = 8
			DO
				EMReadScreen pben_type, 2, pben_row, 24
				IF pben_type = "02" THEN 
					EMReadScreen pben_disp, 1, pben_row, 77
					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN 
						closing_message = closing_message & vbCr & "* Household member " & person & " appears to have pending, appealing, or eligible SSI benefits. Please review for ABAWD and SNAP E&T exemption."
						EXIT DO
					ELSE
						pben_row = pben_row + 1
					END IF
				ELSE
					pben_row = pben_row + 1
				END IF
			LOOP UNTIL pben_row = 14
		END IF
	END IF
NEXT
			
'>>>>>>>>>>PREG
CALL navigate_to_MAXIS_screen("STAT", "PREG")
FOR EACH person IN HH_member_array
	IF person <> "" THEN 
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_PREG, 1, 2, 78
		EMReadScreen preg_end_dt, 8, 12, 53
		IF num_of_PREG <> "0" AND preg_end_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to have active pregnancy. Please review for ABAWD exemption."
	END IF
NEXT
			
'>>>>>>>>>>PROG
CALL navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen cash1_status, 4, 6, 74
EMReadScreen cash2_status, 4, 7, 74
IF cash1_status = "ACTV" OR cash2_status = "ACTV" THEN closing_message = closing_message & vbCr & "* Case is active on CASH programs. Please review for ABAWD and SNAP E&T exemption."
			
			
'>>>>>>>>>SCHL/STIN/STEC
CALL navigate_to_MAXIS_screen("STAT", "SCHL")
FOR EACH person IN HH_member_array
	IF person <> "" THEN 
		CALL write_value_and_transmit(person, 20, 76)
		EMReadScreen num_of_SCHL, 1, 2, 78
		IF num_of_SCHL = "1" THEN 
			EMReadScreen school_status, 1, 6, 40
			IF school_status <> "N" THEN closing_message = closing_message & vbCr & "* Household member " & person & " appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions."
		ELSE
			EMWriteScreen "STIN", 20, 71
			CALL write_value_and_transmit(person, 20, 76)
			EMReadScreen num_of_STIN, 1, 2, 78
			IF num_of_STIN = "1" THEN 
				STIN_row = 8
				DO
					EMReadScreen cov_thru, 5, STIN_row, 67
					IF cov_thru <> "__ __" THEN 
						cov_thru = replace(cov_thru, " ", "/01/")
						cov_thru = DateAdd("M", 1, cov_thru)
						cov_thru = DateAdd("D", -1, cov_thru)
						IF DateDiff("D", date, cov_thru) > 0 THEN 
							closing_message = closing_message & vbCr & "* Household member " & person & " appears to have active student income. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
							EXIT DO
						ELSE
							STIN_row = STIN_row + 1
							IF STIN_row = 18 THEN 
								PF20
								STIN_row = 8
								EMReadScreen last_page, 21, 24, 2
								IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
							END IF
						END IF
					ELSE
						EXIT DO
					END IF
				LOOP 
			ELSE
				EMWriteScreen "STEC", 20, 71
				CALL write_value_and_transmit(person, 20, 76)
				EMReadScreen num_of_STEC, 1, 2, 78
				IF num_of_STEC = "1" THEN 
					STEC_row = 8
					DO
						EMReadScreen stec_thru, 5, STEC_row, 48
						IF stec_thru <> "__ __" THEN 
							stec_thru = replace(stec_thru, " ", "/01/")
							stec_thru = DateAdd("M", 1, stec_thru)
							stec_thru = DateAdd("D", -1, stec_thru)
							IF DateDiff("D", date, stec_thru) > 0 THEN 
								closing_message = closing_message & vbCr & "* Household member " & person & " appears to have active student expenses. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
								EXIT DO
							ELSE
								STEC_row = STEC_row + 1
								IF STEC_row = 17 THEN 
									PF20
									STEC_row = 8
									EMReadScreen last_page, 21, 24, 2
									IF last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
								END IF
							END IF
						ELSE
							EXIT DO
						END IF
					LOOP
				END IF
			END IF	
		END IF
	END IF
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
NEXT

household_persons = ""
pers_count = 0

FOR EACH person IN HH_member_array
	IF person <> "" THEN 
		IF pers_count = uBound(HH_member_array) THEN 
			IF pers_count = 0 THEN 
				household_persons = household_persons & person
			ELSE
				household_persons = household_persons & "and " & person
			END IF
		ELSE
			household_persons = household_persons & person & ", "
			pers_count = pers_count + 1
		END IF
	END IF 
NEXT

IF closing_message = "" THEN 
	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "It appears there are no missed exemptions for ABAWD or SNAP E&T in MAXIS for this case. The script has checked EATS, MEMB, DISA, JOBS, BUSI, RBIC, UNEA, PREG, PROG, PBEN, SCHL, STIN, and STEC for member(s) " & household_persons & "." & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
ELSE
	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "The script has checked for ABAWD and SNAP E&T exemptions coded in MAXIS for member(s) " & household_persons & "." & vbCr & closing_message & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
END IF

STATS_counter = STATS_counter - 1		'Removing one instance from the STATS Counter as it started with one at the beginning
script_end_procedure(closing_message)
