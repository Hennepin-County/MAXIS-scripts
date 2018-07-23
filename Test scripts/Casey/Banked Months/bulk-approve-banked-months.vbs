'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - Approve Banked Months.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 0			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================
run_locally = TRUE
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
call changelog_update("07/11/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================
function excel_open_pw(file_url, visible_status, alerts_status, ObjExcel, objWorkbook, my_password)
'--- This function opens a specific excel file.
'~~~~~ file_url: name of the file
'~~~~~ visable_status: set to either TRUE (visible) or FALSE (not-visible)
'~~~~~ alerts_status: set to either TRUE (show alerts) or FALSE (suppress alerts)
'~~~~~ ObjExcel: leave as 'objExcel'
'~~~~~ objWorkbook: leave as 'objWorkbook'
'===== Keywords: MAXIS, PRISM, MMIS, Excel
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url,,,, my_password) 'Opens an excel file from a specific URL
    ''(file.Path,,,, "mypassword",,,,,,,,,,)
	objExcel.DisplayAlerts = alerts_status
end function

Function review_ABAWD_FSET_exemptions(person_ref_nbr, possible_exemption)
'--- This function screens for ABAWD/FSET exemptions for SNAP.
'===== Keywords: MAXIS, ABAWD, FSET, exemption, SNAP
    CALL check_for_MAXIS(False)

    possible_exemption = FALSE
    closing_message = ""

    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	IF person_ref_nbr <> "" THEN
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen cl_age, 2, 8, 76
		IF cl_age = "  " THEN cl_age = 0
		cl_age = cl_age * 1
		IF cl_age < 18 OR cl_age >= 50 THEN closing_message = closing_message & vbCr & "* M" & person_ref_nbr & ": Appears to have exemption. Age = " & cl_age & "."
	END IF


    CALL navigate_to_MAXIS_screen("STAT", "DISA")
	disa_status = false
	IF person_ref_nbr <> "" THEN
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen num_of_DISA, 1, 2, 78
		IF num_of_DISA <> "0" THEN
			EMReadScreen disa_end_dt, 10, 6, 69
			disa_end_dt = replace(disa_end_dt, " ", "/")
			EMReadScreen cert_end_dt, 10, 7, 69
			cert_end_dt = replace(cert_end_dt, " ", "/")
			IF IsDate(disa_end_dt) = True THEN
				IF DateDiff("D", date, disa_end_dt) > 0 THEN
					closing_message = closing_message & vbCr & "* M" & person_ref_nbr & ": Appears to have disability exemption. DISA end date = " & disa_end_dt & "."
					disa_status = True
				END IF
			ELSE
				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
					closing_message = closing_message & vbCr & "* M" & person_ref_nbr & ": Appears to have disability exemption. DISA has no end date."
					disa_status = True
				END IF
			END IF
			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
				IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* M" & person_ref_nbr & ": Appears to have disability exemption. DISA Certification end date = " & cert_end_dt & "."
			ELSE
				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
					EMReadScreen cert_begin_dt, 8, 7, 47
					IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person_ref_nbr & ": Appears to have disability exemption. DISA certification has no end date."
				END IF
			END IF
		END IF
	END IF

    '>>>>>>>>>>>> EATS GROUP - PLACEHOLDER - WORKING HERE
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
    			find_memb01 = InStr(eats_group, person_ref_nbr)
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
    					IF cl_age = "  " THEN cl_age = 0
    						cl_age = cl_age * 1
    						IF cl_age =< 17 THEN
    							closing_message = closing_message & vbCr & "* M" & person & ": May have exemption for minor child caretaker. Household member " & eats_pers & " is minor. Please review for accuracy."
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
    							closing_message = closing_message & vbCr & "* M" & person & ": MAY have an exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
    							disa_status = TRUE
    						END IF
    					ELSEIF IsDate(disa_end_dt) = False THEN
    						IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
    							closing_message = closing_message & vbCr & "* M" & person & " : MAY have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
    							disa_status = true
    						END IF
    					END IF
    					IF IsDate(cert_end_dt) = True AND disa_status = False THEN
    						IF DateDiff("D", date, cert_end_dt) > 0 THEN closing_message = closing_message & vbCr & "* M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification end date = " & cert_end_dt & "."
    					ELSE
    						IF (cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999") THEN
    							EMReadScreen cert_begin_dt, 8, 7, 47
    							IF cert_begin_dt <> "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification has no end date."
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
    		prospective_hours = 0

    		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
    		EMWritescreen person, 20, 76
    		EMWritescreen "01", 20, 79				'ensures that we start at 1st job
    		transmit
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
    					prospective_hours = prospective_hours + prosp_hrs
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
    						'added seperate incremental variable to account for multiple jobs
    						prospective_hours = prospective_hours + prosp_hrs
    					END IF
    				END IF
    				transmit		'to exit PIC
    				EMReadScreen JOBS_panel_current, 1, 2, 73
    				'looping until all the jobs panels are calculated
    				If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
    			Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
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
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_RBIC, 1, 2, 78
    		IF num_of_RBIC <> "0" THEN closing_message = closing_message & vbCr & "* M" & person & ": Has RBIC panel. Please review for ABAWD and/or SNAP E&T exemption."
    		IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
    			closing_message = closing_message & vbCr & "* M" & person & ": Appears to be working 30 hours/wk (regardless of wage level) or earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions."
    		ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
    			closing_message = closing_message & vbCr & "* M" & person & ": Appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption and SNAP E&T exemptions."
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
    						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
    					END IF
    				ELSE
    					IF unea_end_dt = "__/__/__" THEN
    						IF unea_type = "14" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
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
    			    IF pben_type = "12" THEN		'UI pending'
    					EMReadScreen pben_disp, 1, pben_row, 77
    					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
    						closing_message = closing_message & vbCr & "* M" & person & ": Appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption."
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
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_PREG, 1, 2, 78
            EMReadScreen preg_due_dt, 8, 10, 53
            preg_due_dt = replace(preg_due_dt, " ", "/")
    		EMReadScreen preg_end_dt, 8, 12, 53

    		IF num_of_PREG <> "0" THen
                If preg_due_dt <> "__/__/__" Then
                    If DateDiff("d", date, preg_due_dt) > 0 AND preg_end_dt = "__ __ __" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active pregnancy. Please review for ABAWD exemption."
                    If DateDiff("d", date, preg_due_dt) < 0 Then closing_message = closing_message & vbCr & "* M" & person & ": Appears to have an overdue pregnancy, person may meet a minor child exemption. Contact client."
                End If
            End If
        END IF
    NEXT

    '>>>>>>>>>>PROG
    CALL navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash1_status, 4, 6, 74
    EMReadScreen cash2_status, 4, 7, 74
    IF cash1_status = "ACTV" OR cash2_status = "ACTV" THEN closing_message = closing_message & vbCr & "* Case is active on CASH programs. Please review for ABAWD and SNAP E&T exemption."

    '>>>>>>>>>>ADDR
    CALL navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen homeless_code, 1, 10, 43
    EmReadscreen addr_line_01, 16, 6, 43

    IF homeless_code = "Y" or addr_line_01 = "GENERAL DELIVERY" THEN closing_message = closing_message & vbCr & "* Client is claiming homelessness. If client has barriers to employment, they could meet the 'Unfit for Employment' exemption. Exemption began 05/2018."

    '>>>>>>>>>SCHL/STIN/STEC
    FOR EACH person IN HH_member_array
    	IF person <> "" THEN
            CALL navigate_to_MAXIS_screen("STAT", "SCHL")
    		CALL write_value_and_transmit(person, 20, 76)
    		EMReadScreen num_of_SCHL, 1, 2, 78
    		IF num_of_SCHL = "1" THEN
    			EMReadScreen school_status, 1, 6, 40
    			IF school_status <> "N" THEN closing_message = closing_message & vbCr & "* M" & person & ": Appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions."
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
    							closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active student income. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
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
    								closing_message = closing_message & vbCr & "* M" & person & ": Appears to have active student expenses. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
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
    	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "It appears there are NO missed exemptions for ABAWD or SNAP E&T in MAXIS for this case. The script has checked ADDR, EATS, MEMB, DISA, JOBS, BUSI, RBIC, UNEA, PREG, PROG, PBEN, SCHL, STIN, and STEC for member(s) " & household_persons & "." & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    ELSE
    	closing_message = "*** NOTICE!!! ***" & vbCr & vbCr & "The script has checked for ABAWD and SNAP E&T exemptions coded in MAXIS for member(s) " & household_persons & "." & vbCr & closing_message & vbCr & vbCr & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    END IF

    'Displaying the results...now with added MsgBox bling.
    'vbSystemModal will keep the results in the foreground.
    MsgBox closing_message, vbInformation + vbSystemModal, "ABAWD/FSET Exemption Check -- Results"

    STATS_counter = STATS_counter - 1		'Removing one instance from the STATS Counter as it started with one at the beginning
End Function
'===========================================================================================================================

'NEED A SCRIPT THAT WILL OPERATE OFF OF THE DAIL - PEPR (this is from a list generated with BULK-Dail)
    'FS ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS
    'Review case for possible ABAWD exemptions, 2nd set, then finally Banked Months
        'if banked - add to working list
    'Gathers MEMB number and which months are banked
    'Updates ABAWD tracking record and WREG
    'approves case
    'CASE NOTES

'NEED A SCRIPT TO ASSES AND UPDATE A WORKING EXCEL
    'There is a BOBI list of all clients on SNAP
    'It should compare to a working list and for any not on the working list
    'For any not on the list, asses for potential banked months cases

'NEED A SCRIPT TO REVIEW ALL THE CASES on the BANKED MONTHS LIST'

'THOUGHTS
'Use case notes instead of person notes to document used Banked Months as that way we are using a form people are more comfortable with.

'CONSTANTS=================================================================================================================

'THE COLUMNS IN THE WORKING EXCEL
Const case_nbr_col      = 1
Const memb_nrb_col      = 2
Const last_name_col     = 3
Const first_name_col    = 4
Const notes_col         = 5
Const first_mo_col      = 6
Const scnd_mo_col       = 7
Const third_mo_col      = 8
Const fourth_mo_col     = 9
Const fifth_mo_col      = 10
Const sixth_mo_col      = 11
Const svnth_mo_col      = 12
Const eighth_mo_col     = 13
Const ninth_mo_col      = 14
Const curr_mo_stat_col  = 15


'CONSTANTS FOR ARRAYS

Const case_nbr          =  0
Const clt_excel_row     =  1
Const memb_ref_nbr      =  2
Const clt_last_name     =  3
Const clt_first_name    =  4
Const clt_notes         =  5
Const clt_mo_one        =  6
Const clt_mo_two        =  7
Const clt_mo_three      =  8
Const clt_mo_four       =  9
Const clt_mo_five       = 10
Const clt_mo_six        = 11
Const clt_mo_svn        = 12
Const clt_mo_eight      = 13
Const clt_mo_nine       = 14
Const clt_curr_mo_stat  = 14
Const case_errors       = 15

'==========================================================================================================================


'THE SCRIPT================================================================================================================

'Connects to BlueZone
EMConnect ""

Dim BANKED_MONTHS_CASES_ARRAY ()
ReDim BANKED_MONTHS_CASES_ARRAY (case_errors, 0)

'Initial Dialog will have worker select which option is going to be run at this time
    'Assess Banked Month cases from DAIL PEPR List
    'Review monthly BOBI report of all SNAP clients
    'Review of Banked Months cases
    'Approve ongoing Banked Month Cases
    'HAVE DEVELOPER MODE

'IF NOT in Developer Mode, check to be sure we are in production

BeginDialog Dialog1, 0, 0, 181, 80, "Dialog"
  DropListBox 15, 35, 160, 45, "Ongoing Banked Months Cases", process_option
  ButtonGroup ButtonPressed
    OkButton 70, 60, 50, 15
    CancelButton 125, 60, 50, 15
  Text 10, 10, 170, 10, "Script to assess and review Banked Months cases."
EndDialog

dialog Dialog1
cancel_confirmation

excel_row_to_start = "2"

BeginDialog Dialog1, 0, 0, 116, 95, "Dialog"
  EditBox 75, 25, 30, 15, stop_time
  EditBox 75, 50, 30, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
  Text 10, 10, 95, 10, "Script will run "
  Text 15, 30, 50, 10, "Hours"
  Text 15, 55, 50, 10, "Excel to start"
EndDialog

dialog Dialog1

excel_row_to_start = excel_row_to_start * 1

'making stop time a number
stop_time = FormatNumber(stop_time, 2,          0,                 0,                      0)
                        'number     dec places  leading 0 - FALSE    neg nbr in () - FALSE   use deliminator(comma) - FALSE
stop_time = stop_time * 60 * 60     'tunring hours to seconds

end_time = timer + stop_time        'timer is the number of seconds from 12:00 AM so we need to add the hours to run to the time to determine at what point the script should exit the loop


'DAIL PEPR Option
    'Dialog to select the Excel list that has the DAILs
    'add all to an array
    'Compare the array to the Working list
        'add to working list if not already there

'BOBI Report Option
    'Check each person on the BOBI list in MAXIS
        'exclude clients with obvious exclusions (?? age)
        'should we actually check MAXIS to see if it is coded correctly?
    'add each to the array
    'compare the array to the working list
        'if not already on the list, check WREG for 30/13


'Review of cases
    'Open the working Excel sheet
    'Have worker confirm the correct sheet opened
    'Read all the cases from the spreadsheet and add to an array

    'Check CASE CURR to see if case and person are still active SNAP
        'If closed need to review if the closure was correct
    'Check WREG
        'Confirm case is coded as 30/13
        'Confirm ABAWD months have been used
    'GET Code from UTILITIES - COUNTED ABAWD MONTHS - need to confirm that counted months are correct
    'GET Code from ACTIONS - ABAWD FSET EXEMPTION CHECK and run it on every SNAP month to check the counting
    'Update MAXIS panels/WREG/ABAWD Tracking Record as determined by other runs
        'may need to do person search to see if there was SNAP on another case that caused the counted month
        'If any month is confusing then use code from NOTES - ABAWD TRACKING RECORD to coordinate
        'MAY need dialog for worker to confirm confusing months
    'Need to check ECF - create a dialog to allow worker to review ECF information
    '????'

'Approve ongoing cases
    'Open the working Excel sheet
    'Have worker confirm the correct sheet opened
    'Read all the cases from the spreadsheet and add to an array

    'Read PROG and CASE/PERS to confirm client is still active SNAP on this case
    'Check for possible exemption in STAT
    'Review Case Notes to see if there are any case notes that need to be assessed
        'Have a series of case notes that can be ignored
        'Look just to the last BM case note
        'Have a dialog for the worker to review the case notes if anything appears indicating a change may have happened
        'Worker can confirm that the BM coding is correct or adjust in the dialog
    'Go to WREG
    'Check tracker to see if any ABAWD months have fallen off of the 36 month look back period
    'Update WREG with any information found
        'If exempt - update exemption coding
        'If still BM ensure coding is 30/13 and update the BM counter
    'Review case and update other STAT panels if eneeded (JOBS dates)
    'Review ELIG and approve
    'Update Excel

If process_option = "Ongoing Banked Months Cases" Then
    'working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\On Demand Waiver\Files for testing new application rewrite\Working Excel.xlsx"
    working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Master banked months list.xlsx"     'THIS IS THE REAL ONE

    'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
    call excel_open_pw(working_excel_file_path, True, True, ObjExcel, objWorkbook, "BM")

    ObjExcel.Worksheets("Ongoing banked months").Activate
    ObjExcel.SendKeys "{ENTER}"

    list_row = excel_row_to_start
    the_case = 0
    Do
        ReDim Preserve BANKED_MONTHS_CASES_ARRAY(case_errors, the_case)
        BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case)           = trim(ObjExcel.Cells(list_row, case_nbr_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_excel_row, the_case)      = list_row
        BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)       = trim(ObjExcel.Cells(list_row, memb_nrb_col).Value)

        BANKED_MONTHS_CASES_ARRAY(clt_last_name, the_case)      = trim(ObjExcel.Cells(list_row, last_name_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_first_name, the_case)     = trim(ObjExcel.Cells(list_row, first_name_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case)          = trim(ObjExcel.Cells(list_row, notes_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_one, the_case)         = trim(ObjExcel.Cells(list_row, first_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_two, the_case)         = trim(ObjExcel.Cells(list_row, scnd_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_three, the_case)       = trim(ObjExcel.Cells(list_row, third_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_four, the_case)        = trim(ObjExcel.Cells(list_row, fourth_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_five, the_case)        = trim(ObjExcel.Cells(list_row, fifth_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_six, the_case)         = trim(ObjExcel.Cells(list_row, sixth_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_svn, the_case)         = trim(ObjExcel.Cells(list_row, svnth_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_eight, the_case)       = trim(ObjExcel.Cells(list_row, eighth_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_mo_nine, the_case)        = trim(ObjExcel.Cells(list_row, ninth_mo_col).Value)
        BANKED_MONTHS_CASES_ARRAY(clt_curr_mo_stat, the_case)   = trim(ObjExcel.Cells(list_row, curr_mo_stat_col).Value)

        list_row = list_row + 1
        the_case = the_case + 1

    Loop Until trim(ObjExcel.Cells(list_row, case_nbr_col).Value) = ""

    function set_lastest_banked_month(date_variable, month_mo, month_yr, boo_var)
        month_mo = left(date_variable, 2)
        month_yr = right(date_variable, 2)
        If month_mo <> CM_plus_1_mo Then boo_var = FALSE
        If month_yr <> CM_plus_1_yr Then boo_var = FALSE
    end function

    For the_case = 0 to UBOUND(BANKED_MONTHS_CASES_ARRAY, 2)

        list_row = BANKED_MONTHS_CASES_ARRAY(clt_excel_row, the_case)
        MAXIS_case_number = BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case)
        BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) = Right("00"&BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 2)
        ObjExcel.Cells(list_row, memb_nrb_col) = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)

        For month_indicator = clt_mo_one to clt_mo_nine
            Call back_to_SELF
            month_tracked = FALSE
            abawd_status = ""
            fset_wreg_status = ""

            If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) <> "" Then                          'if the spreadsheet is already full
                MAXIS_footer_month = left(BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case), 2)
                MAXIS_footer_year = right(BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case), 2)

                month_tracked = TRUE
            Else
                first_of_footer_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year     'this is set from the last month that was entered in the spreadsheet
                next_month = DateAdd("m", 1, first_of_footer_month)

                MAXIS_footer_month = DatePart("m", next_month)
                MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)

                MAXIS_footer_year = DatePart("yyyy", next_month)
                MAXIS_footer_year = right(MAXIS_footer_year, 2)
            End If

            Call navigate_to_MAXIS_screen("CASE", "PERS")
            pers_row = 10
            clt_SNAP_status = ""
            Do
                EMReadScreen pers_ref_numb, 2, pers_row, 3
                If pers_ref_numb = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) Then
                    EMReadScreen clt_SNAP_status, 1, pers_row, 54
                    Exit Do
                Else
                    pers_row = pers_row + 3
                    If pers_row = 19 Then
                        PF8
                        pers_row = 10
                    End If
                End If
            Loop until pers_ref_numb = "  "

            If clt_SNAP_status = "A" Then
                Call navigate_to_MAXIS_screen("STAT", "WREG")

                EMWriteScreen BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 20, 76
                transmit

                EMReadScreen fset_wreg_status, 2, 8, 50
                EMReadScreen abawd_status, 2, 13, 50

                month_tracker_nbr = month_indicator - 5
                update_WREG = FALSE

                If fset_wreg_status = "30" AND abawd_status = "13" Then
                    If MAXIS_footer_month = CM_mo AND MAXIS_footer_year = CM_yr Then update_WREG = TRUE

                    If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then update_WREG = TRUE

                    If update_WREG = TRUE Then

                        'Need to be sure that there isn't a new ABAWD month available - maybe another column with the counted months on the ongoing banked months cases
                        'Need to review case for possible exemption months - code from exemption finder
                        Call review_ABAWD_FSET_exemptions

                        ' PF9
                        ' EMWriteScreen month_tracker_nbr, 14, 50
                        ' transmit
                        ' EMWriteScreen "BGTX", 20, 71
                        ' transmit

                        'Write TIKL or something to identify cases to be approved and noted.
                        'IDEA - write a new column in to Excel for cases needing approval in months
                        'IDEA - write a process that will send a case through background and stop with a dialog to allow for manual approval.
                        BANKED_MONTHS_CASES_ARRAY(notes_col, the_case) = BANKED_MONTHS_CASES_ARRAY(notes_col, the_case) & " ~ Approve SNAP for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "~"
                        'TODO need to find a casenoting solution for these months
                    End If
                End If

                If month_tracked  = FALSE Then
                    BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year
                End If

            Else
                If month_tracked = TRUE Then

                    BeginDialog Dialog1, 0, 0, 191, 110, "Dialog"
                      ButtonGroup ButtonPressed
                        PushButton 15, 75, 160, 15, "Yes - remove the month from the Master List", yes_remove_month_btn
                        PushButton 15, 95, 160, 10, "No - keep the month - case will be updated", no_keep_btn
                      Text 30, 10, 130, 15, "It appears that for the month MM/YY the Member 01 was not active on SNAP."
                      Text 30, 35, 130, 15, "This month has been tracked on the Banked Month master list."
                      Text 10, 60, 180, 10, "Should the month be removed from the tracking sheet?"
                    EndDialog

                    dialog Dialog1

                    If ButtonPressed = yes_remove_month_btn Then BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = ""

                Else
                    BANKED_MONTHS_CASES_ARRAY(case_errors, the_case) = "STALE"
                    BANKED_MONTHS_CASES_ARRAY(notes_col, the_case) = BANKED_MONTHS_CASES_ARRAY(notes_col, the_case) & "  ~Client is not active SNAP in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " ~  "
                    'MsgBox "STALE"
                End If
            End If
            'MsgBox "Column " & ObjExcel.Cells(1, month_indicator) & " for tracking says - " & BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) & vbNewLine & "For the month of " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for the case: " & MAXIS_case_number & vbNewLine & "Member " & BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) & " is " & clt_SNAP_status & "." & vbNewLine & "WREG is FSET - " & fset_wreg_status & " | ABAWD - " & abawd_status
            If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then Exit For
        Next

        '************************************************************************************'
        ' banked_months_tracked = TRUE
        '
        ' If BANKED_MONTHS_CASES_ARRAY(clt_mo_nine, the_case) <> "" Then
        '     MsgBox "NINE"
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_eight, the_case) <> "" Then
        '     MsgBox "EIGHT"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_eight, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_eight
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_svn, the_case) <> "" Then
        '     MsgBox "SEVEN"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_svn, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_svn
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_six, the_case) <> "" Then
        '     MsgBox "SIX"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_six, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_six
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_five, the_case) <> "" Then
        '     MsgBox "FIVE"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_five, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_five
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_four, the_case) <> "" Then
        '     MsgBox "FOUR"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_four, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_four
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_three, the_case) <> "" Then
        '     MsgBox "THREE"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_three, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_three
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_two, the_case) <> "" Then
        '     MsgBox "TWO"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_two, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_two
        ' ElseIf BANKED_MONTHS_CASES_ARRAY(clt_mo_one, the_case) <> "" Then
        '     MsgBox "ONE"
        '     Call set_lastest_banked_month(BANKED_MONTHS_CASES_ARRAY(clt_mo_one, the_case), banked_month_mo, banked_month_yr, banked_months_tracked)
        '     last_month_tracked = clt_mo_one
        ' End If
        ' MsgBox banked_month_mo & "/" & banked_month_yr
        ' 'look at each month from the last approved to CM plus 1 to review if clt is active and wreg is 30/13 then track
        ' 'if CM or CM plus one, do new approval will WREG counter updated'
        ' If banked_months_tracked = FALSE Then
        '     Call back_to_SELF
        '     first_of_footer_month = banked_month_mo & "/01/" & banked_month_yr
        '     next_month = Datedd("m", 1, first_of_footer_month)
        '
        '     MAXIS_footer_month = DatePart("m", next_month)
        '     MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)
        '
        '     MAXIS_footer_year = DatePart("yyyy", next_month)
        '     MAXIS_footer_year = right(&MAXIS_footer_year, 2)
        '
        '     For col_to_update = (last_month_tracked+1) to clt_mo_nine
        '         Call navigate_to_MAXIS_screen("CASE", "PERS")
        '         pers_row = 10
        '         clt_SNAP_status = ""
        '         Do
        '             EMReadScreen pers_ref_numb, 2, pers_row, 3
        '             If pers_ref_numb = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) Then
        '                 EMReadScreen clt_SNAP_status, 1, pers_row, 54
        '                 Exit Do
        '             Else
        '                 pers_row = pers_row + 3
        '                 If pers_row = 19 Then
        '                     PF8
        '                     pers_row = 10
        '                 End If
        '             End If
        '         Loop until pers_ref_numb = "  "
        '
        '         If clt_SNAP_status <> "A" Then
        '             Call navigate_to_MAXIS_screen("STAT", "WREG")
        '
        '             EMWriteScreen BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 20, 76
        '             transmit
        '
        '
        '         first_of_footer_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
        '         next_month = Datedd("m", 1, first_of_footer_month)
        '
        '         MAXIS_footer_month = DatePart("m", next_month)
        '         MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)
        '
        '         MAXIS_footer_year = DatePart("yyyy", next_month)
        '         MAXIS_footer_year = right(&MAXIS_footer_year, 2)
        '
        '     Next
        ' End If

        ' ObjExcel.Cells(list_row, case_nbr_col).Value        = BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case)
        ' ObjExcel.Cells(list_row, memb_nrb_col).Value        = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)
        ' ObjExcel.Cells(list_row, last_name_col).Value       = BANKED_MONTHS_CASES_ARRAY(clt_last_name, the_case)
        ' ObjExcel.Cells(list_row, first_name_col).Value      = BANKED_MONTHS_CASES_ARRAY(clt_first_name, the_case)

        ObjExcel.Cells(list_row, notes_col).Value           = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case)
        ObjExcel.Cells(list_row, first_mo_col).Value        = BANKED_MONTHS_CASES_ARRAY(clt_mo_one, the_case)
        ObjExcel.Cells(list_row, scnd_mo_col).Value         = BANKED_MONTHS_CASES_ARRAY(clt_mo_two, the_case)
        ObjExcel.Cells(list_row, third_mo_col).Value        = BANKED_MONTHS_CASES_ARRAY(clt_mo_three, the_case)
        ObjExcel.Cells(list_row, fourth_mo_col).Value       = BANKED_MONTHS_CASES_ARRAY(clt_mo_four, the_case)
        ObjExcel.Cells(list_row, fifth_mo_col).Value        = BANKED_MONTHS_CASES_ARRAY(clt_mo_five, the_case)
        ObjExcel.Cells(list_row, sixth_mo_col).Value        = BANKED_MONTHS_CASES_ARRAY(clt_mo_six, the_case)
        ObjExcel.Cells(list_row, svnth_mo_col).Value        = BANKED_MONTHS_CASES_ARRAY(clt_mo_svn, the_case)
        ObjExcel.Cells(list_row, eighth_mo_col).Value       = BANKED_MONTHS_CASES_ARRAY(clt_mo_eight, the_case)
        ObjExcel.Cells(list_row, ninth_mo_col).Value        = BANKED_MONTHS_CASES_ARRAY(clt_mo_nine, the_case)
        ObjExcel.Cells(list_row, curr_mo_stat_col).Value    = BANKED_MONTHS_CASES_ARRAY(clt_curr_mo_stat, the_case)

        If timer > end_time Then
            end_msg = "Success! Script has run for " & stop_time/60/60 & " hours and has finished for the time being."
            Exit For
        End If
    Next

End If
'
'NEED another spreadsheet for all cases that WERE banked months cases but are no longer - so that we can save the case information

script_end_procedure(end_msg)
