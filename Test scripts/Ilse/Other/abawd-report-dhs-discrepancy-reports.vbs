worker_county_code = "x127"
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - ABAWD REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each CASE
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
call changelog_update("06/17/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function BULK_ABAWD_FSET_exemption_finder(possible_exemptions)
'--- This function screens for ABAWD/FSET exemptions for SNAP.
'===== Keywords: MAXIS, ABAWD, FSET, exemption, SNAP
'----------------------------------------------------------------------------------------------------Determining the EATS Household
    possible_exemptions = ""

    CALL navigate_to_MAXIS_screen("STAT", "EATS")
    eats_group_members = ""
    memb_found = True
    EMReadScreen all_eat_together, 1, 4, 72

    IF all_eat_together = "_" THEN
        eats_group_members = "01" & "," 'single member HH's
    ELSEIF all_eat_together = "Y" THEN
    'HH's where all members eat together
        eats_row = 5
        DO
            EMReadScreen eats_eats_pers, 2, eats_row, 3
            eats_eats_pers = replace(eats_eats_pers, " ", "")
            IF eats_eats_pers <> "" THEN
                eats_group_members = eats_group_members & eats_eats_pers & ","
                eats_row = eats_row + 1
            END IF
        LOOP UNTIL eats_eats_pers = ""
    ELSEIF all_eat_together = "N" THEN
    'multiple eats HH cases - we are only caring about the 1st eats group that contains MEMB 01.
        eats_row = 13
        DO
            EMReadScreen eats_group, 38, eats_row, 39
            find_memb01 = InStr(eats_group, eats_pers)
            IF find_memb01 = 0 THEN
                eats_row = eats_row + 1
                IF eats_row = 18 THEN
                    memb_found = False
                    EXIT DO
                END IF
            END IF
        LOOP UNTIL find_memb01 <> 0

        'Gathering the eats group members
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
        eats_group_members = trim(eats_group_members)
        eats_group_members = split(eats_group_members, ",")

        IF all_eat_together <> "_" THEN
            CALL write_value_and_transmit("MEMB", 20, 71)
            FOR EACH eats_pers IN eats_group_members
                IF eats_pers <> "" AND eats_pers <> eats_pers THEN
                    CALL write_value_and_transmit(eats_pers, 20, 76)
                    EMReadScreen cl_age, 2, 8, 76
                    IF cl_age = "  " THEN cl_age = 0
                        cl_age = cl_age * 1
                        IF cl_age =< 17 THEN
                            possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": May have exemption for minor child caretaker. Household member " & eats_pers & " is minor. Please review for accuracy."
                        END IF
                END IF
            NEXT
        END IF

        CALL navigate_to_MAXIS_screen("STAT", "MEMB")
        FOR EACH eats_pers IN eats_group_members
        	IF eats_pers <> "" THEN
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen cl_age, 2, 8, 76
        		IF cl_age = "  " THEN cl_age = 0
        		cl_age = cl_age * 1
        		IF cl_age < 18 OR cl_age >= 50 THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have exemption. Age = " & cl_age & "."
        	END IF
        NEXT

        CALL navigate_to_MAXIS_screen("STAT", "DISA")
        FOR EACH eats_pers IN eats_group_members
        	disa_status = false
        	IF eats_pers <> "" THEN
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen num_of_DISA, 1, 2, 78
        		IF num_of_DISA <> "0" THEN
        			EMReadScreen disa_end_dt, 10, 6, 69
        			disa_end_dt = replace(disa_end_dt, " ", "/")
        			EMReadScreen cert_end_dt, 10, 7, 69
        			cert_end_dt = replace(cert_end_dt, " ", "/")
        			IF IsDate(disa_end_dt) = True THEN
        				IF DateDiff("D", date, disa_end_dt) > 0 THEN
        					possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have disability exemption. DISA end date = " & disa_end_dt & "."
        					disa_status = True
        				END IF
        			ELSE
        				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
        					possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have disability exemption. DISA has no end date."
        					disa_status = True
        				END IF
        			END IF
        			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
        				IF DateDiff("D", date, cert_end_dt) > 0 THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have disability exemption. DISA Certification end date = " & cert_end_dt & "."
        			ELSE
        				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
        					EMReadScreen cert_begin_dt, 8, 7, 47
        					IF cert_begin_dt <> "__ __ __" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have disability exemption. DISA certification has no end date."
        				END IF
        			END IF
        		END IF
        	END IF
        NEXT


        CALL write_value_and_transmit("DISA", 20, 71)
        FOR EACH disa_pers IN eats_group_members
        	disa_status = false
        	IF disa_pers <> "" AND disa_pers <> eats_pers THEN
        		CALL write_value_and_transmit(disa_pers, 20, 76)
        		EMReadScreen num_of_DISA, 1, 2, 78
        		IF num_of_DISA <> "0" THEN
        			EMReadScreen disa_end_dt, 10, 6, 69
        			disa_end_dt = replace(disa_end_dt, " ", "/")
        			EMReadScreen cert_end_dt, 10, 7, 69
        			cert_end_dt = replace(cert_end_dt, " ", "/")
        			IF IsDate(disa_end_dt) = True THEN
        				IF DateDiff("D", date, disa_end_dt) > 0 THEN
        					possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": MAY have an exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
        					disa_status = TRUE
        				END IF
        			ELSEIF IsDate(disa_end_dt) = False THEN
        				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
        					possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & " : MAY have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
        					disa_status = true
        				END IF
        			END IF
        			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
        				IF DateDiff("D", date, cert_end_dt) > 0 THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification end date = " & cert_end_dt & "."
        			ELSE
        				IF (cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999") THEN
        					EMReadScreen cert_begin_dt, 8, 7, 47
        					IF cert_begin_dt <> "__ __ __" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification has no end date."
        				END IF
        			END IF
        		END IF
        	END IF
        NEXT

        '>>>>>>>>>>>>>>EARNED INCOME
        FOR EACH eats_pers IN eats_group_members
        	IF eats_pers <> "" THEN
        		prosp_inc = 0
        		prosp_hrs = 0
        		prospective_hours = 0

        		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
        		EMWritescreen eats_pers, 20, 76
        		EMWritescreen "01", 20, 79				'ensures that we start at 1st job
        		transmit
        		EMReadScreen num_of_JOBS, 1, 2, 78
        		IF num_of_JOBS <> "0" THEN
        			DO
        			 	EMReadScreen jobs_end_dt, 8, 9, 49
        				EMReadScreen cont_end_dt, 8, 9, 73
        				IF jobs_end_dt = "__ __ __" THEN
        					CALL write_value_and_transmit("X", 19, 38)     'Entering the PIC
        					EMReadScreen prosp_monthly, 8, 18, 56
        					prosp_monthly = trim(prosp_monthly)
        					IF prosp_monthly = "" THEN prosp_monthly = 0
        					prosp_inc = prosp_inc + prosp_monthly
        					EMReadScreen prosp_hrs, 8, 16, 50
        					IF prosp_hrs = "        " THEN prosp_hrs = 0
        					prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
        					EMReadScreen pay_freq, 1, 5, 64
        					IF pay_freq = "1" THEN
        						prosp_hrs = prosp_hrs
        					ELSEIF pay_freq = "2" THEN
        						prosp_hrs = (2 * prosp_hrs)
        					ELSEIF pay_freq = "3" THEN
        						prosp_hrs = (2.15 * prosp_hrs)
        					ELSEIF pay_freq = "4" THEN
        						prosp_hrs = (4.3 * prosp_hrs)
        					END IF
                            transmit		'to exit PIC
        					prospective_hours = prospective_hours + prosp_hrs
        				ELSE
        					jobs_end_dt = replace(jobs_end_dt, " ", "/")
        					IF DateDiff("D", date, jobs_end_dt) > 0 THEN
        						'Going into the PIC for a job with an end date in the future
        						CALL write_value_and_transmit("X", 19, 38)        'Entering the PIC
        						EMReadScreen prosp_monthly, 8, 18, 56
        						prosp_monthly = trim(prosp_monthly)
        						IF prosp_monthly = "" THEN prosp_monthly = 0
        						prosp_inc = prosp_inc + prosp_monthly
        						EMReadScreen prosp_hrs, 8, 16, 50
        						IF prosp_hrs = "        " THEN prosp_hrs = 0
        						prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
        						EMReadScreen pay_freq, 1, 5, 64
        						IF pay_freq = "1" THEN
        							prosp_hrs = prosp_hrs
        						ELSEIF pay_freq = "2" THEN
        							prosp_hrs = (2 * prosp_hrs)
        						ELSEIF pay_freq = "3" THEN
        							prosp_hrs = (2.15 * prosp_hrs)
        						ELSEIF pay_freq = "4" THEN
        							prosp_hrs = (4.3 * prosp_hrs)
        						END IF
                                transmit		'to exit PIC
        						'added seperate incremental variable to account for multiple jobs
        						prospective_hours = prospective_hours + prosp_hrs
        					END IF
        				END IF

        				EMReadScreen JOBS_panel_current, 1, 2, 73
        				'looping until all the jobs panels are calculated
        				If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
        			Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
        		END IF

        		EMWriteScreen "BUSI", 20, 71
        		CALL write_value_and_transmit(eats_pers, 20, 76)
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
        						prospective_hours = prospective_hours + busi_hrs
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
        						prospective_hours = prospective_hours + busi_hrs
        					END IF
        				END IF
        				transmit
        				EMReadScreen enter_a_valid, 13, 24, 2
        			LOOP UNTIL enter_a_valid = "ENTER A VALID"
        		END IF

        		EMWriteScreen "RBIC", 20, 71
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen num_of_RBIC, 1, 2, 78
        		IF num_of_RBIC <> "0" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Has RBIC panel. Please review for ABAWD and/or SNAP E&T exemption."
        		IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
        			possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to be working 30 hours/wk (regardless of wage level) or earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions."
        		ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
        			possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption and SNAP E&T exemptions."
        		END IF
        	END IF
        NEXT

        '>>>>>>>>>>>>UNEA
        CALL navigate_to_MAXIS_screen("STAT", "UNEA")
        FOR EACH eats_pers IN eats_group_members
        	IF eats_pers <> "" THEN
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen num_of_UNEA, 1, 2, 78
        		IF num_of_UNEA <> "0" THEN
        			DO
        				EMReadScreen unea_type, 2, 5, 37
        				EMReadScreen unea_end_dt, 8, 7, 68
        				unea_end_dt = replace(unea_end_dt, " ", "/")
        				IF IsDate(unea_end_dt) = True THEN
        					IF DateDiff("D", date, unea_end_dt) > 0 THEN
        						IF unea_type = "14" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
        					END IF
        				ELSE
        					IF unea_end_dt = "__/__/__" THEN
        						IF unea_type = "14" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
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
        FOR EACH eats_pers IN eats_group_members
        	IF eats_pers <> "" THEN
        		EMWriteScreen "PBEN", 20, 71
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen num_of_PBEN, 1, 2, 78
        		IF num_of_PBEN <> "0" THEN
        			pben_row = 8
        			DO
        			    IF pben_type = "12" THEN		'UI pending'
        					EMReadScreen pben_disp, 1, pben_row, 77
        					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
        						possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption."
        						EXIT DO
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
        FOR EACH eats_pers IN eats_group_members
        	IF eats_pers <> "" THEN
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen num_of_PREG, 1, 2, 78
                EMReadScreen preg_due_dt, 8, 10, 53
                preg_due_dt = replace(preg_due_dt, " ", "/")
        		EMReadScreen preg_end_dt, 8, 12, 53

        		IF num_of_PREG <> "0" THen
                    If preg_due_dt <> "__/__/__" Then
                        If DateDiff("d", date, preg_due_dt) > 0 AND preg_end_dt = "__ __ __" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have active pregnancy. Please review for ABAWD exemption."
                        If DateDiff("d", date, preg_due_dt) < 0 Then possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have an overdue pregnancy, eats_pers may meet a minor child exemption. Contact client."
                    End If
                End If
            END IF
        NEXT

        '>>>>>>>>>>PROG
        CALL navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen cash1_status, 4, 6, 74
        EMReadScreen cash2_status, 4, 7, 74
        IF cash1_status = "ACTV" OR cash2_status = "ACTV" THEN possible_exemptions = possible_exemptions & vbCr & "* Case is active on CASH programs. Please review for ABAWD and SNAP E&T exemption."

        '>>>>>>>>>>ADDR
        CALL navigate_to_MAXIS_screen("STAT", "ADDR")
        EMReadScreen homeless_code, 1, 10, 43
        EmReadscreen addr_line_01, 16, 6, 43

        IF homeless_code = "Y" or addr_line_01 = "GENERAL DELIVERY" THEN possible_exemptions = possible_exemptions & vbCr & "* Client is claiming homelessness. If client has barriers to employment, they could meet the 'Unfit for Employment' exemption. Exemption began 05/2018."

        '>>>>>>>>>SCHL/STIN/STEC
        FOR EACH eats_pers IN eats_group_members
        	IF eats_pers <> "" THEN
                CALL navigate_to_MAXIS_screen("STAT", "SCHL")
        		CALL write_value_and_transmit(eats_pers, 20, 76)
        		EMReadScreen num_of_SCHL, 1, 2, 78
        		IF num_of_SCHL = "1" THEN
        			EMReadScreen school_status, 1, 6, 40
        			IF school_status <> "N" THEN possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions."
        		ELSE
        			EMWriteScreen "STIN", 20, 71
        			CALL write_value_and_transmit(eats_pers, 20, 76)
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
        							possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have active student income. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
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
        				CALL write_value_and_transmit(eats_pers, 20, 76)
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
        								possible_exemptions = possible_exemptions & vbCr & "* M" & eats_pers & ": Appears to have active student expenses. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
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

        'household_eats_perss = ""
        'pers_count = 0

        'FOR EACH eats_pers IN eats_group_members
        '	IF eats_pers <> "" THEN
        '		IF pers_count = uBound(HH_member_array) THEN
        '			IF pers_count = 0 THEN
        '				household_eats_perss = household_eats_perss & eats_pers
        '			ELSE
        '				household_eats_perss = household_eats_perss & "and " & eats_pers
        '			END IF
        '		ELSE
        '			household_eats_perss = household_eats_perss & eats_pers & ", "
        '			pers_count = pers_count + 1
        '		END IF
        '	END IF
        'NEXT

        IF possible_exemptions = "" THEN possible_exemptions = "No missed exemptions."
    End if
End Function



'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
worker_county_code = "X127"
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

file_selection_path = "C:\Users\ilfe001\OneDrive - Hennepin County\Desktop\SNAP Work\ABAWD WREG Clean Up Report 3.2023.xlsx"
worksheet_name = "Exemptions Indicated"

'column constants
pmi_col         =  3
case_number_col =  2
fset_col        = 4
abawd_col       = 5
memb_numb_col   = 15
snap_status_col = 16
notes_col       = 17
case_active_col = 18
exemptions_col  = 19

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "BULK - ABAWD REPORT"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
  EditBox 85, 95, 20, 15, MAXIS_footer_month
  EditBox 110, 95, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  GroupBox 10, 5, 250, 85, "Using this script:"
  Text 20, 100, 65, 10, "Footer month/year:"
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  Text 20, 20, 235, 25, "This script should be used when a list of SNAP cases wtih member numbers are provided by BOBI to gather ABAWD, FSET and Banked Months information."
EndDialog

Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
         Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

objExcel.worksheets(worksheet_name).Activate

'Setting the Excel rows with variables
ObjExcel.Cells(1, memb_numb_col).Value = "Member #"
ObjExcel.Cells(1, snap_status_col).Value = "SNAP Status"
ObjExcel.Cells(1, notes_col).Value = "Notes"
ObjExcel.Cells(1, case_active_col).Value = "Case Active"
ObjExcel.Cells(1, exemptions_col).Value = "Possible Exemptions"

FOR i = 1 to 22		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
'    If objWorkSheet.Name = "10-20" then sheet_list = sheet_list & objWorkSheet.Name & ","
'Next

'For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
'    If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All cases" and objWorkSheet.Name <> "Data" then sheet_list = sheet_list & objWorkSheet.Name & ","
'Next
'
'sheet_list = trim(sheet_list)  'trims excess spaces of sheet_list
'If right(sheet_list, 1) = "," THEN sheet_list = left(sheet_list, len(sheet_list) - 1) 'trimming off last comma
'array_of_sheets = split(sheet_list, ",")   'Creating new array
'
'For each excel_sheet in array_of_sheets
''    objExcel.worksheets(excel_sheet).Activate 'Activates the applicable worksheet

    'MAXIS_footer_month = left(excel_sheet, 2)
    'MAXIS_footer_year = right(excel_sheet, 2)
    Call MAXIS_footer_month_confirmation

    excel_row = 1000

    Do
    	PMI_number = trim(ObjExcel.Cells(excel_row, pmi_col).Value)

        MAXIS_case_number = ObjExcel.Cells(excel_row, case_number_col).Value
    	MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do

        Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
        EmReadscreen self_screen, 4, 2, 50
        EmReadscreen self_error, 60, 24, 2
        If is_this_priv = True then
            ObjExcel.Cells(excel_row, notes_col).Value = "Privliged case"
        Elseif (is_this_priv = False and self_screen = "SELF") then
            ObjExcel.Cells(excel_row, notes_col).Value = trim(self_error)
        Else
            Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
            ObjExcel.Cells(excel_row, snap_status_col).Value = snap_case
            ObjExcel.Cells(excel_row, case_active_col).Value = case_active

            EmReadscreen county_code, 4, 21, 14 'reading from CASE/CURR
            If county_code <> UCASE(worker_county_code) then
                ObjExcel.Cells(excel_row, notes_col).Value = "Out-of-county Case"
            Else
                Call navigate_to_MAXIS_screen("STAT", "MEMB")
                Do
                    EmReadscreen memb_panel_PMI, 8, 4, 46
                    'memb_panel_PMI = right ("00000000" & trim(memb_panel_PMI), 8)
                    If trim(memb_panel_PMI) = PMI_number then
                        EmReadscreen member_number, 2, 4, 33
                        Exit do
                    Else
                        transmit
                        EmReadscreen end_of_membs_message, 5, 24, 2
                    End if
                Loop until end_of_membs_message = "ENTER"

                If trim(member_number) = "" then
                    ObjExcel.Cells(excel_row, notes_col).Value = "Unable to find member on case"
                Else
    	            Call navigate_to_MAXIS_screen("STAT", "WREG")
                    Call write_value_and_transmit(member_number, 20, 76)

    	            EMReadScreen FSET_code, 2, 8, 50
    	            EMReadScreen ABAWD_code, 2, 13, 50

                    Call BULK_ABAWD_FSET_exemption_finder(possible_exemptions)

                    ObjExcel.Cells(excel_row, memb_numb_col).Value = member_number                      'writing in the member number with initial 0 trimmed.
                    ObjExcel.Cells(excel_row, fset_col).Value = replace(FSET_code, "_", "")
    	            ObjExcel.Cells(excel_row, abawd_col).Value = replace(ABAWD_code, "_", "")
                    ObjExcel.Cells(excel_row, exemptions_col).Value = trim(possible_exemptions)
                End if
            End if
        End if
        STATS_counter = STATS_counter + 1
        excel_row = excel_row + 1
    Loop until ObjExcel.Cells(excel_row, 2).Value = ""
'Next

FOR i = 1 to 21		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")
