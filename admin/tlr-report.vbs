'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TLR REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                      'manual run time in seconds
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

Function ABAWD_Tracking_Record(abawd_counted_months, member_number, MAXIS_footer_month)
    EMReadScreen wreg_panel, 4, 2, 48
    If wreg_panel <> "WREG" then Call navigate_to_MAXIS_screen("STAT","WREG")		'navigates to stat/wreg
    EMReadScreen wreg_memb, 2, 4, 33
    If wreg_memb <> member_number then CALL write_value_and_transmit(member_number, 20, 76)
    Call write_value_and_transmit("X", 13, 57) 'Pulls up the WREG tracker'
    EMWaitReady 0, 0
    EMReadscreen tracking_record_check, 15, 4, 40  		'adds cases to the rejection list if the ABAWD tracking record cannot be accessed.
    EMWaitReady 0,0
    If tracking_record_check <> "Tracking Record" then
		report_notes = report_notes & "Error accessing ATR. "
    ELSE
        TLR_fixed_clock_mo = "01" 'fixed clock dates for all recipients 
	    TLR_fixed_clock_yr = "23"
    
	    bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
        bene_yr_row = 10
        abawd_counted_months = 0					'declares the variables values at 0
        month_count = 0
    
        DO
        	'establishing variables for specific ABAWD counted month dates
        	If bene_mo_col = "19" then counted_date_month = "01"
        	If bene_mo_col = "23" then counted_date_month = "02"
        	If bene_mo_col = "27" then counted_date_month = "03"
        	If bene_mo_col = "31" then counted_date_month = "04"
        	If bene_mo_col = "35" then counted_date_month = "05"
        	If bene_mo_col = "39" then counted_date_month = "06"
        	If bene_mo_col = "43" then counted_date_month = "07"
        	If bene_mo_col = "47" then counted_date_month = "08"
        	If bene_mo_col = "51" then counted_date_month = "09"
        	If bene_mo_col = "55" then counted_date_month = "10"
        	If bene_mo_col = "59" then counted_date_month = "11"
        	If bene_mo_col = "63" then counted_date_month = "12"
        	'counted date year: this is found on rows 7-10. Row 11 is current year plus one, so this will be exclude this list.
        	If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", date), 2)
        	If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, date)), 2)
        	If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, date)), 2)
        	If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)
            abawd_counted_months_string = counted_date_month & "/" & counted_date_year
    
        	'reading to see if a month is counted month or not
        	EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
        	'counting and checking for counted ABAWD months
        	IF is_counted_month = "X" or is_counted_month = "M" THEN
        		EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
        		abawd_counted_months = abawd_counted_months + 1				'adding counted months
        	END IF
        	bene_mo_col = bene_mo_col - 4		're-establishing search once the end of the row is reached
        	IF bene_mo_col = 15 THEN
        		bene_yr_row = bene_yr_row - 1
        		bene_mo_col = 63
        	END IF
        			
	    'used to loop until count was 36 due to person based look back period. Now fixed clock starts 01/23 for all members. 
        LOOP until (counted_date_month = TLR_fixed_clock_mo AND counted_date_year = TLR_fixed_clock_yr)
        PF3	' to exit tracking record 
    End if 
End Function

Function BULK_ABAWD_FSET_exemption_finder()
'excluding matching grant and participating in CD treatment due to non-MAXIS indicators.
'excluding armed forces participation dur to non-MAXIS indicators. 
'----------------------------------------------------------------------------------------------------Determining the EATS Household
    'default strings and counts
	verified_wreg = ""
	verified_abawd = ""
	eats_HH_count = 0
	possible_exemptions = ""

    CALL navigate_to_MAXIS_screen("STAT", "EATS")
    eats_group_members = ""
    memb_found = True
    EMReadScreen all_eat_together, 1, 4, 72

    IF all_eat_together = "_" THEN
        eats_group_members = "01" & "," 'single member HH's
		eats_HH_count = 1
    ELSEIF all_eat_together = "Y" THEN
    'HH's where all members eat together
        eats_row = 5
        DO
            EMReadScreen eats_pers, 2, eats_row, 3
            eats_pers = replace(eats_pers, " ", "")
            IF eats_pers <> "" THEN
                eats_group_members = eats_group_members & eats_pers & ","
				eats_HH_count = eats_HH_count  + 1
                eats_row = eats_row + 1
            END IF
        LOOP UNTIL eats_pers = ""
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
				eats_HH_count = eats_HH_count  + 1
            END IF
        LOOP UNTIL eats_group = "__"
    END IF

	ObjExcel.Cells(excel_row, eats_HH_col).Value = eats_HH_count

	'Case-based determination
    '----------------------------------------------------------------------------------------------------14 – ES Compliant While Receiving MFIP
	'----------------------------------------------------------------------------------------------------20 – ES Compliant While Receiving DWP
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
	If mfip_case = True then verified_wreg = verified_wreg & "14" & "|"
	If DWP_case = True then verified_wreg = verified_wreg & "20" & "|"
    
	ObjExcel.Cells(excel_row, snap_status_col).Value = snap_status

	'----------------------------------------------------------------------------------------------------17 – Receiving RCA
	'Case-based determination -- Looking for RCA information while still on CASE/CURR	
	row = 1                                            
    col = 1
    EMSearch "RCA:", row, col
    If row <> 0 Then
        EMReadScreen rca_status, 9, row, col + 5
        rca_status = trim(rca_status)
		rca_status = rca_status
        If rca_status = "ACTIVE" or rca_status = "APP CLOSE" or rca_status = "APP OPEN" Then
            rca_case = TRUE
			verified_wreg = verified_wreg & "17" & "|"
        End If
	End if 
	
	'----------------------------------------------------------------------------------------------------'temp coding - Foster care on 18th 
	''<<<<<<<<<<PROG for Foster care 
	'Person-based evaluation
    CALL navigate_to_MAXIS_screen("STAT", "PROG")
	EmReadScreen IV-E_prog, 8, 11, 33 
	EMReadScreen IV-E_status, 4, 11, 74
	If trim(IV-E_prog) = "__ __ __" or IV-E_prog = 0 then 
		foster_care = False 
	else 
		If Trim(IV-E_status) <> "DENY" then 
			foster_care = True
		else 
			foster_care = False 
		End if 
	End if

	Call HCRE_panel_bypass	'making sure we don't get stuck 

	'Case-based determination
    IF memb_found = True THEN
		If SNAP_status <> "INACTIVE" then
            eats_group_members = trim(eats_group_members)
            eats_group_members = split(eats_group_members, ",")

		    child_under_six = False 	'defaulting to False
		    child_under_18 = False		'defaulting to False
			adult_HH_count = 0

            IF all_eat_together <> "_" THEN
                CALL navigate_to_MAXIS_screen("STAT", "MEMB")
                FOR EACH eats_pers IN eats_group_members
                    IF trim(eats_pers) <> "" THEN
                        CALL write_value_and_transmit(eats_pers, 20, 76)
                        EMReadScreen cl_age, 2, 8, 76
                        cl_age = trim(cl_age)
                        IF cl_age = "" THEN cl_age = 0
                        cl_age = cl_age * 1
						If cl_age < 6 then child_under_six = True
                        IF cl_age =< 17 THEN
							child_under_18 = True
		    			Else
							adult_HH_count = adult_HH_count + 1
		    			End if
                    END IF
                NEXT
            END IF

		    '----------------------------------------------------------------------------------------------------21 – Child < 18 Living in the SNAP Unit
 		    If child_under_18 = True then verified_wreg = verified_wreg & "21" & "|"

			'----------------------------------------------------------------------------------------------------08 – Responsible for care of child <6 years old
			If child_under_six = True then
				If adult_HH_count = 1 then
					verified_wreg = verified_wreg & "08" & "|"
				Else
					possible_exemptions = possible_exemptions & vbcr & "Child under 6 is in the SNAP Household. "
				End if
			End if

		    'person-based determination
			age_50 = False
            age_53_54 = False 
            age_53_54_counted = False 'temporary coding to support. Effective 10/1/24 53-54 YO's starting being TLR's after their next renewal
            
            CALL navigate_to_MAXIS_screen("STAT", "MEMB")
            CALL write_value_and_transmit(member_number, 20, 76)
            EMReadScreen cl_age, 2, 8, 76
            cl_age = trim(cl_age)
		    EMReadScreen age_verif_code, 2, 8, 68
            IF cl_age = "" THEN cl_age = 0
            cl_age = cl_age * 1

		    '----------------------------------------------------------------------------------------------------07 – Age 16-17, Living W/Pare/Crgvr
		    If cl_age = 16 or cl_age = 17 then
		    	EMReadScreen age_verif_code, 2, 8, 68
		    	If age_verif_code <> "NO" then
		    		verified_wreg = verified_wreg & "07" & "|"
		    	End if
		    End if

		    '----------------------------------------------------------------------------------------------------06 – Under age 16
		    If cl_age < 16 then
		    	If age_verif_code <> "NO" then
		    		verified_wreg = verified_wreg & "06" & "|"
		    	End if
		    End if
		    '----------------------------------------------------------------------------------------------------'16 – 53-59 Years Old
		    If cl_age => 53 then
		    	If cl_age < 60 then
		    		If age_verif_code <> "NO" then
		    			verified_wreg = verified_wreg & "16" & "|"
		    		End if
		    	End if
		    End if
		    '----------------------------------------------------------------------------------------------------'05 - Age 60 or older
		    If cl_age => 60 then
		    If age_verif_code <> "NO" then
		    	verified_wreg = verified_wreg & "05" & "|"
		    	End if
		    End if

			'----------------------------------------------------------------------------------------------------special temp handling for 50-52 year olds later on based on age_50 = True
			If cl_age = 50 or _
				cl_age = 51 or _ 
				cl_age = 52 then 
				age_50 = True
			End if 

            If cl_age = 53 or _
                cl_age = 54 then
                age_53_54 = True
            End if 

			'----------------------------------------------------------------------------------------------------possible exemption for foster care members under 24 YO. 
			If cl_age < 24 then 
				If foster_care = True then possible_exemptions = possible_exemptions & vbcr & "Member is under 24 & may have been in foster case on 18th birthday. Review case. "
			End if 
			
			'<<<<<<<<<<DISA
			'Case-based evaluation
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
            				IF DateDiff("D", ABAWD_eval_date, disa_end_dt) > 0 THEN
								disa_status = True
            					If eats_pers <> member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " - DISA end date = " & disa_end_dt & ". "
            				END IF
            			ELSE
            				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
								disa_status = True
								If eats_pers <> member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " -DISA has no end date. "
            				END IF
            			END IF
            			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
            				IF DateDiff("D", ABAWD_eval_date, cert_end_dt) > 0 THEN
								If eats_pers <> member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " - " & cert_end_dt & ". "
							End if
						ELSE
            				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
            					EMReadScreen cert_begin_dt, 8, 7, 47
            					IF cert_begin_dt <> "__ __ __" THEN
									If eats_pers <> member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " -DISA certification has no end date. "
								End if
							END IF
            			END IF
            		END IF
            	END IF
            NEXT

			'Person based evaluation
            'Still in DISA
            CALL write_value_and_transmit(member_number, 20, 76)
			verified_disa = False
			disa_status = False
            EMReadScreen num_of_DISA, 1, 2, 78

			IF num_of_DISA <> "0" THEN
            	EMReadScreen disa_end_dt, 10, 6, 69
            	disa_end_dt = replace(disa_end_dt, " ", "/")
            	EMReadScreen cert_end_dt, 10, 7, 69
            	cert_end_dt = replace(cert_end_dt, " ", "/")
            	IF IsDate(disa_end_dt) = True THEN
            		IF DateDiff("D", ABAWD_eval_date, disa_end_dt) > 0 THEN disa_status = True
            	ELSE
            		IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN disa_status = True
            	END IF
            	IF IsDate(cert_end_dt) = True AND disa_status = False THEN
            		IF DateDiff("D", ABAWD_eval_date, cert_end_dt) > 0 THEN disa_status = True
				ELSE
            		IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
            			EMReadScreen cert_begin_dt, 8, 7, 47
            			IF cert_begin_dt <> "__ __ __" THEN disa_status = True
					End if
				End if
			END IF

            If disa_status = True then
                row = 11
                Do
                    EmReadscreen prog_disa_code, 2, row, 59
                    If prog_disa_code <> "__" then
                        EmReadscreen prog_disa_verif, 1, row, 69
                        If prog_disa_verif <> "N" then
                            If row = 11 or row = 13 then
                                verified_disa = True
                                exit do
                            Else
                                If prog_disa_verif = "7" then
                                    verified_disa = False
                                Else
                                    verified_disa = True
                                    exit do
                                End if
                            End if
                        End if
                    End if
                    row = row + 1
                Loop until row = 14

				If verified_disa = True then verified_wreg = verified_wreg & "03" & "|"
			End if

            '>>>>>>>>>>>>>>EARNED INCOME
		    'Person-based determination for Earned Income
            prosp_inc = 0
            prosp_hrs = 0
            prospective_hours = 0
            CALL navigate_to_MAXIS_screen("STAT", "JOBS")
            EMWritescreen member_number, 20, 76
		    Call write_value_and_transmit("01", 20, 79)				'ensures that we start at 1st job
            EMReadScreen num_of_JOBS, 1, 2, 78
            IF num_of_JOBS <> "0" THEN
            	DO
            	 	EMReadScreen jobs_end_dt, 8, 9, 49
            		EMReadScreen cont_end_dt, 8, 9, 73
            		IF jobs_end_dt = "__ __ __" THEN
		    			EMReadScreen jobs_verif_code, 1, 6, 34
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
            			IF DateDiff("D", ABAWD_eval_date, jobs_end_dt) > 0 THEN
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
            				'added separate incremental variable to account for multiple jobs
            				prospective_hours = prospective_hours + prosp_hrs
            			END IF
            		END IF
            		EMReadScreen JOBS_panel_current, 1, 2, 73
            		'looping until all the jobs panels are calculated
            		If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
            	Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
            END IF

		    'Person-based determination
            EMWriteScreen "BUSI", 20, 71
            CALL write_value_and_transmit(member_number, 20, 76)
            EMReadScreen num_of_BUSI, 1, 2, 78
            IF num_of_BUSI <> "0" THEN
            	DO
            		EMReadScreen busi_end_dt, 8, 5, 72
            		busi_end_dt = replace(busi_end_dt, " ", "/")
            		IF IsDate(busi_end_dt) = True THEN
		    			Call write_value_and_transmit("X", 6, 26) 'entering gross income calculation pop-up
		    			EMReadScreen busi_verif_code, 1, 11, 73
		    			PF3 'to exit pop up
            			IF DateDiff("D", ABAWD_eval_date, busi_end_dt) > 0 THEN
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

		    'Person based since very unlikely to be case based at this point.
            EMWriteScreen "RBIC", 20, 71
            CALL write_value_and_transmit(member_number, 20, 76)
            EMReadScreen num_of_RBIC, 1, 2, 78
            IF num_of_RBIC <> "0" then ObjExcel.Cells(excel_row, notes_col).Value = "Actually found an RBIC."
	
            IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
		    	If jobs_verif_code <> "N" or jobs_verif_code <> "N" then
		    		If busi_verif_code <> "_" or busi_verif_code <> "N" then
		    			verified_wreg = verified_wreg & "09" & "|"
		    		End if
		    	End if
            ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
		    	If jobs_verif_code <> "N" or jobs_verif_code <> "N" then
		    		If busi_verif_code <> "_" or busi_verif_code <> "N" then
		    			verified_abawd = verified_wreg & "06"
		    		End if
		    	End if
            END IF

            '>>>>>>>>>>>>UNEA
		    '----------------------------------------------------------------------------------------------------'03 – Unfit for Employment
		    'Person-based determination
            CALL write_value_and_transmit(member_number, 20, 76)
            EMReadScreen num_of_UNEA, 1, 2, 78
            IF num_of_UNEA <> "0" THEN
            	DO
            		EMReadScreen unea_type, 2, 5, 37
            		EMReadScreen unea_end_dt, 8, 7, 68
            		unea_end_dt = replace(unea_end_dt, " ", "/")
            		IF IsDate(unea_end_dt) = True THEN
            			IF DateDiff("D", ABAWD_eval_date, unea_end_dt) > 0  or unea_end_dt = "__/__/__" THEN
            				IF unea_type = "11" then
		    					EmReadScreen VA_verif_code, 1, 5, 65
		    					If VA_verif_code <> "N" then
		    						verified_wreg = verified_wreg & "03" & "|"
		    						Exit do
		    					Else
		    						If eats_pers = member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to have VA disability benefits. "
		    					End if
		    				End if
            			END IF
            		END IF
            		transmit
            		EMReadScreen enter_a_valid, 13, 24, 2
            	LOOP UNTIL enter_a_valid = "ENTER A VALID"
            END IF

		    '----------------------------------------------------------------------------------------------------'11 – Rcvg UI or Work Compliant While UI Pending
		    'Person-based determination
            uc_unea = False 
            CALL write_value_and_transmit(member_number, 20, 76)
            EMReadScreen num_of_UNEA, 1, 2, 78
            IF num_of_UNEA <> "0" THEN
            	DO
            		EMReadScreen unea_type, 2, 5, 37
            		EMReadScreen unea_end_dt, 8, 7, 68
            		unea_end_dt = replace(unea_end_dt, " ", "/")
            		IF IsDate(unea_end_dt) = True THEN
            			IF DateDiff("D", ABAWD_eval_date, unea_end_dt) > 0  or unea_end_dt = "__/__/__" THEN
            				IF unea_type = "14" then
		    					EmReadScreen UC_verif_code, 1, 5, 65
		    					If UC_verif_code <> "N" then
                                    uc_unea = True 
		    						verified_wreg = verified_wreg & "11" & "|"
		    						Exit do
		    					Else
		    						If eats_pers = member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to have active unemployment benefits. "
		    					End if
		    				End if
            			END IF
            		END IF
            		transmit
            		EMReadScreen enter_a_valid, 13, 24, 2
            	LOOP UNTIL enter_a_valid = "ENTER A VALID"
            END IF

		    '----------------------------------------------------------------------------------------------------'11 – Rcvg UI or Work Compliant While UI Pending
            '>>>>>>>>>PBEN
		    'Person based determination
            CALL navigate_to_MAXIS_screen("STAT", "PBEN")
		    Call write_value_and_transmit(member_number, 20, 76)
		    EMReadScreen num_of_PBEN, 1, 2, 78
            IF num_of_PBEN <> "0" THEN
            	pben_row = 8
            	DO
                    EMreadscreen pben_type, 2, pben_row, 24
                    If pben_type = "__" then exit do
            	    IF pben_type = "12" THEN		'UI pending'
            			EMReadScreen pben_disp, 1, pben_row, 77
            			IF pben_disp = "A" OR pben_disp = "P" THEN
		    				verified_wreg = verified_wreg & "11" & "|"
		    				EXIT DO
                        elseif pben_disp = "E" then 
                            if uc_unea = True then 
                                verified_wreg = verified_wreg & "11" & "|"
                                Exit do
                            Else
                                pben_row = pben_row + 1
                            End if 
            			Else
		    				If eats_pers = member_number then possible_exemptions = possible_exemptions & vbcr & "May have pending, appealing, or eligible Unemployment benefits. "
                            pben_row = pben_row + 1
            			END IF
            		ELSE
            			pben_row = pben_row + 1
            		END IF
            	LOOP UNTIL pben_row = 12
		    End if

		    '----------------------------------------------------------------------------------------------------23 – Pregnant
            '>>>>>>>>>>PREG
		    'Person based determination
            CALL navigate_to_MAXIS_screen("STAT", "PREG")
			Call write_value_and_transmit(member_number, 20, 76)
		    EMReadScreen num_of_PREG, 1, 2, 78
            IF num_of_PREG <> "0" THEN
                EMReadScreen preg_due_dt, 8, 10, 53
                preg_due_dt = replace(preg_due_dt, " ", "/")
            	EMReadScreen preg_end_dt, 8, 12, 53

                If preg_due_dt <> "__/__/__" Then
		    		EMReadscreen preg_verif, 1, 6, 75
                    If DateDiff("d", ABAWD_eval_date, preg_due_dt) >= 0 AND preg_end_dt = "__ __ __" THEN

						If preg_verif = "Y" then
							verified_wreg = verified_wreg & "23" & "|"
						Elseif preg_verif = "N" then 
							verified_wreg = verified_wreg & "23" & "|"
						Elseif preg_verif = "?" then 
							verified_wreg = verified_wreg & "23" & "|"	'expedited coding is fine for the exemption. 
						Else 
							possible_exemptions = possible_exemptions & vbcr & "Appears to have an unverified active pregnancy. "
						End if
					End If
				End if
            End If
		    '----------------------------------------------------------------------------------------------------30/09 - Military Servive
            '>>>>>>>>>>MEMI
		    'Person-based determination
            CALL navigate_to_MAXIS_screen("STAT", "MEMI")
			Call write_value_and_transmit(member_number, 20, 76)
            EMReadScreen military_service_code, 1, 12, 78
            If military_service_code = "Y" then 
                verified_wreg = verified_wreg & "30" & "|" 
            End if 

            '>>>>>>>>>>ADDR
		    'Case based determination
			homeless_exemption = False
            CALL navigate_to_MAXIS_screen("STAT", "ADDR")
            EMReadScreen homeless_code, 1, 10, 43
			EMReadScreen living_situation, 2, 11, 43
            EmReadscreen addr_line_01, 16, 6, 43

            IF homeless_code = "Y" then
				If living_situation = "02" or _
					living_situation = "06" or _							
					living_situation = "07" or _
					living_situation = "08" then 
					verified_wreg = verified_wreg & "03" & "|"
					homeless_exemption = True 
				Else
					possible_exemptions = possible_exemptions & vbcr & "Case's ADDR is coded Y for homeless but living situation doesn't match. "  
				End if 
            Elseif addr_line_01 = "GENERAL DELIVERY" THEN
                possible_exemptions = possible_exemptions & vbcr & "Case's ADDR is General Delivery. "
			Else 
				homeless_exemption = False
            End if

            '>>>>>>>>>SCHL/STIN/STEC
		    'person based determination
		    CALL navigate_to_MAXIS_screen("STAT", "SCHL")
            CALL write_value_and_transmit(member_number, 20, 76)
            EMReadScreen num_of_SCHL, 1, 2, 78
            IF num_of_SCHL = "1" THEN
            	EMReadScreen school_status, 1, 6, 40
                EMReadScreen school_verif, 2, 6, 63
                EMReadScreen SNAP_code, 2, 16, 63
            	IF school_status = "F" or school_status = "H" then
                    If school_verif = "SC" or school_verif = "OT" then
                        If  SNAP_code = "01" or _
                            SNAP_code = "02" or _
                            SNAP_code = "04" or _
                            SNAP_code = "05" or _
                            SNAP_code = "06" or _
                            SNAP_code = "07" or _
                            SNAP_code = "09" or _
                            SNAP_code = "10" then
                            verified_wreg = verified_wreg & "12" & "|"
                        Else
                            If eats_pers = member_number then possible_exemptions = possible_exemptions & vbcr & "Appears to be in school w/ unverified school status. "
                        End if
                    End if
                End if
		    End if

            IF possible_exemptions = "" THEN possible_exemptions = "No other potential exemptions. "
		End if

	    'filter the list here for best_wreg_code
	    If trim(verified_wreg) = "" then
	    	best_wreg_code = "30"
            If verified_abawd = "" then
	    		best_abawd_code = "10"
	    	Else
	    		best_abawd_code = verified_abawd 'this should only be 06 for now but maybe more later
	    	End if
	    Elseif len(verified_wreg) = 3 then
	    	best_wreg_code = replace(verified_wreg, "|", "")
        Else 
            wreg_hierarchy = array("03","04","05","06","07","08","09","10","11","12","13","14","20","15","16","21","17","23","30")
            for each code in wreg_hierarchy
                If instr(verified_wreg, code) then
                    best_wreg_code = code
                    exit for
                End if
            next
	    End if
        'Use this determination for 53-54 YO's
        If age_53_54 = True then 
            If len(verified_wreg) > 2 then
                wreg_hierarchy = array("03","04","05","06","07","08","09","10","11","12","13","14","20","15","21","17","23","16","30")  'Code 16 moved to be a lower exemption 
                for each code in wreg_hierarchy
                    If instr(verified_wreg, code) then
                        best_wreg_code = code
                        exit for
                    End if
                next
            End if         
        End if 
    
	    If trim(best_abawd_code) = "" then 
            If best_wreg_code = "03" or _
	    	    best_wreg_code = "04" or _
	    	    best_wreg_code = "05" or _
	    	    best_wreg_code = "06" or _
	    	    best_wreg_code = "07" or _
	    	    best_wreg_code = "08" or _
	    	    best_wreg_code = "09" or _ 
	    	    best_wreg_code = "10" or _
	    	    best_wreg_code = "11" or _
	    	    best_wreg_code = "12" or _
	    	    best_wreg_code = "13" or _
	    	    best_wreg_code = "14" or _
	    	    best_wreg_code = "20" then
	    	        best_abawd_code = "01"
	        End if
	        If best_wreg_code = "15" then best_abawd_code = "02"
	        If best_wreg_code = "16" then best_abawd_code = "03"
	        If best_wreg_code = "21" then best_abawd_code = "04"
	        If best_wreg_code = "17" then best_abawd_code = "12"
	        If best_wreg_code = "23" then best_abawd_code = "05"
            If best_wreg_code = "30" then best_abawd_code = "09" 'This is for military Service folks only since that is the only thing we can read for in MAXIS to determine the verified_wreg code. Otherwise anyone who is TLR the verified_wreg is "".
        End If 

		'Adding in handling for the next SNAP renewal - these don't need to be assigned if renewal is next month. Just them getting updated is enough. 
		Call navigate_to_MAXIS_screen("STAT", "REVW")
		EMReadScreen next_revw_mo, 2, 9, 57
		EMReadScreen next_revw_yr, 2, 9, 63
		next_SNAP_revw = next_revw_mo & "/" & next_revw_yr
		next_month = CM_plus_1_mo & "/" & CM_plus_1_yr
		
        If next_SNAP_revw = next_month then report_notes = report_notes & "SNAP Review Next Month. "   

        'Checking/managing 53-54 yo's who don't start counting until their next review. 
        If (age_53_54 = True and best_wreg_code = "16") then 
            Call navigate_to_MAXIS_screen("REPT", "ACTV")
            EMReadScreen actv_case_number, 8, 7, 12
            actv_case_number = trim(actv_case_number)
            actv_case_number = right(00000000 & actv_case_number, 8)
            
            If MAXIS_case_number = trim(actv_case_number) then 
                EMReadScreen next_revw_date, 8, 7, 42
                next_revw_date = replace(next_revw_date, " ", "/")
                report_notes = report_notes & next_revw_date & " - 53-54 YO becomes TLR. "
            End if 
        End If      
	    If best_wreg_code = "30" or age_50 = True then Call ABAWD_Tracking_Record(abawd_counted_months, member_number, MAXIS_footer_month)
        updates_needed = True
    
		'----------------------------------------------------------------------------------------------------Age 50 - 52 WREG and ABAWD Tracking Record Handling 
		age_50_workaround = False
        manual_code = "F"  'manual code for exemption cases  
        If age_50 = True then
			'changing codes per temp policy 
			best_wreg_code = "16"
			best_abawd_code = "03"
            age_50_workaround = True 
            manual_code = "M" 
        End if  

	    Call navigate_to_MAXIS_screen("STAT", "WREG")
        Call write_value_and_transmit(member_number, 20, 76)
        PF9
		EMWriteScreen best_wreg_code, 8, 50
		EMWriteScreen best_abawd_code, 13, 50
		If best_wreg_code = "30" then
		    EmWriteScreen "N", 8, 80
		Else
		    EMWriteScreen "_", 8, 80
		End if

        'Updating the ATR if the codes are already not updated for the CM      
        ATR_updates = array("D",manual_code)
        For each update_code in ATR_updates
           Call write_value_and_transmit("X", 13, 57) 'Pulls up the WREG tracker'              
            bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))      'col to search starts at 15, increased by 4 for each footer month
            bene_yr_row = 10
            EMReadScreen ATR_code, 1, bene_yr_row, bene_mo_col
            'This bit will only update to the manual codes if the month isn't already reflecting that. 
            If manual_code = "F" then 
                If ATR_code = "E" or ATR_code = "F" then
                    exit for 'F and E are exmept
                Else 
                    Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
                End if 
            ELSEIF manual_code = "M" then 
                If ATR_code = "X" or ATR_code = "M" then 
                    exit for 'X and M are counted months
                Else 
                    Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
                End if 
            Else 
                Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
            End if 
           PF3 'to go back to WREG/Panel
        Next

        'Count all the ABAWD months
        Call ABAWD_Tracking_Record(abawd_counted_months, member_number, MAXIS_footer_month) 
	    transmit ' to save 
		EMReadscreen orientation_warning, 7, 24, 2 	'reading for orientation date warning message. This message has been casuing me TROUBLE!!
		If orientation_warning = "WARNING" then transmit 
	    PF3 'to save and exit to stat/wrap

	    'case note workaround
        If age_50_workaround = True then 
	        start_a_blank_CASE_NOTE
            Call write_variable_in_CASE_NOTE("--SNAP Time Limited Recipient: Age " & cl_age & "--")	
		    Call write_variable_in_CASE_NOTE("TLR member #" & member_number)
	        Call write_variable_in_CASE_NOTE("---")
	        Call write_variable_in_CASE_NOTE("* Effective 10/23 50-52 year olds are no longer exempt from SNAP time limits due solely to age.")
	        Call write_variable_in_CASE_NOTE("* FSET/ABAWD codes continue to be 16/03 until DHS system updates are in place. ABAWD Tracking record has been updated for this month as a counted month per policy.")
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE(Worker_Signature)
	        PF3
		    ObjExcel.Cells(excel_row, notes_col).Value = cl_age & " year old!"
	    End if

        'script will update the WREG panel for the member if an update
        Call navigate_to_MAXIS_screen("STAT", "WREG")
        Call write_value_and_transmit(member_number, 20, 76)
        PF9
        EMWriteScreen best_wreg_code, 8, 50
        EMWriteScreen best_abawd_code, 13, 50
        If best_wreg_code = "30" then
            EmWriteScreen "N", 8, 80
        Else
            EMWriteScreen "_", 8, 80
        End if
    
        transmit ' to save 
		EMReadscreen orientation_warning, 7, 24, 2 	'reading for orientation date warning message. This message has been casuing me TROUBLE!!
		If orientation_warning = "WARNING" then transmit 
	    PF3 'to save and exit to stat/wrap

	    If homeless_exemption = True then
	        start_a_blank_CASE_NOTE
            Call write_variable_in_CASE_NOTE("--SNAP Time Limited Exempt: Homelessness--")	
	    	Call write_variable_in_CASE_NOTE("---")
	    	Call write_variable_in_CASE_NOTE("* Case is code as homeless on ADDR, and has applicable living situation which exempts this case from SNAP Work Rules and time limits.")
			Call write_variable_in_CASE_NOTE("* FSET/ABAWD codes are 03/01 for members whom meet this exemption.")
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE(Worker_Signature)
	    	PF3
	    End if
    End if 

    'Additional notes for the assignment as to when to give it out. Basically if the approval or data wreg/abawd codes match the best codes they don't need to get updated or reassigned. 
    If updates_needed = True then 
        If snap_status = "ACTIVE" then
            If data_wreg = best_wreg_code then
                If data_abawd = best_abawd_code then
	    			updates_needed = False
                    report_notes = report_notes & "No Updates Needed. "
                End if
            End if
	    Else 
            report_notes = report_notes & "SNAP is " & snap_status & ". "
        End if
    End if 

	ObjExcel.Cells(excel_row, best_WREG_col).Value = best_wreg_code
    ObjExcel.Cells(excel_row, best_abawd_col).Value = best_abawd_code
    ObjExcel.Cells(excel_row, notes_col).Value = report_notes
	ObjExcel.Cells(excel_row, verified_wreg_col).Value = verified_wreg
	ObjExcel.Cells(excel_row, counted_months_col).Value = abawd_counted_months
    ObjExcel.Cells(excel_row, all_exemptions_col).Value = trim(possible_exemptions)
End Function

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
worker_county_code = "X127"
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
ABAWD_eval_date = CM_plus_1_mo & "/01/" & CM_plus_1_yr

'file_selection_path = "C:\Users\ilfe001\OneDrive - Hennepin County\Assignments\" & CM_mo & "-20" & CM_yr & " ABAWD-TLR's.xlsx"
file_selection_path = "C:\Users\ilfe001\OneDrive - Hennepin County\Assignments\11-2024 ABAWD-TLR's.xlsx" 'testing code

'column constants
case_number_col 	= 1		'Col A
pmi_col         	= 2		'Col B
SNAP_status_col	    = 3		'Col C
memb_numb_col   	= 4		'Col D
eats_HH_col			= 5		'Col E
Data_ABAWD_col		= 6		'Col F
Data_wreg_col 		= 7		'Col G
CM_wreg_col		   	= 8		'Col H
CM_abawd_col		= 9		'Col I
best_wreg_col		= 10	'Col J
best_abawd_col		= 11	'Col K
notes_col			= 12	'Col L
                            'Col M - 13 Assignee Name
verified_wreg_col 	= 14	'Col N
counted_months_col	= 15	'Col O
all_exemptions_col	= 16	'Col P

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "ADMIN - TLR REPORT"
  ButtonGroup ButtonPressed
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  GroupBox 10, 5, 250, 85, "Using this script:"
  Text 20, 20, 235, 25, "This script should be used when a list of SNAP recipients with member numbers to assess Time-Limited recipients (TLR's)."
  EditBox 15, 50, 180, 15, file_selection_path
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  Text 20, 100, 65, 10, "Footer month/year:"
  EditBox 85, 95, 20, 15, MAXIS_footer_month
  EditBox 110, 95, 20, 15, MAXIS_footer_year
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
Call MAXIS_footer_month_confirmation

excel_row = 2
Do
    MAXIS_case_number = ""
	MAXIS_case_number = ObjExcel.Cells(excel_row, case_number_col).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do

	PMI_number = trim(ObjExcel.Cells(excel_row, pmi_col).Value)
	PMI_number = right ("00000000" & trim(PMI_number), 8)
	ObjExcel.Cells(excel_row, pmi_col).Value = PMI_number

	data_wreg =  trim(ObjExcel.Cells(excel_row, data_wreg_col).Value)
	data_abawd = trim(ObjExcel.Cells(excel_row, data_abawd_col).Value)

    report_notes = ""

    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
    If is_this_priv = True then
        report_notes = report_notes & "Don't assign - Privliged case. "
    Else
        Call MAXIS_background_check     'needed when more than one member on a case is on a list.
        Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
        EmReadscreen county_code, 4, 21, 14 'reading from CASE/CURR
        If county_code <> UCASE(worker_county_code) then
            report_notes = report_notes & "Don't assign - Out-of-county Case. "
        Else
            Call navigate_to_MAXIS_screen("STAT", "MEMB")
            Do
                EmReadscreen memb_panel_PMI, 8, 4, 46
                memb_panel_PMI = right ("00000000" & trim(memb_panel_PMI), 8)
                If trim(memb_panel_PMI) = PMI_number then
                    EmReadscreen member_number, 2, 4, 33
					ObjExcel.Cells(excel_row, memb_numb_col).Value = member_number
					Exit do
                Else
                    transmit
                    EmReadscreen end_of_membs_message, 5, 24, 2
                End if
            Loop until end_of_membs_message = "ENTER"
            If trim(member_number) = "" then
                report_notes = report_notes = "Unable to find member on case"
            Else
	            Call navigate_to_MAXIS_screen("STAT", "WREG")
                Call write_value_and_transmit(member_number, 20, 76)
	            EMReadScreen FSET_code, 2, 8, 50
	            EMReadScreen ABAWD_code, 2, 13, 50
				ObjExcel.Cells(excel_row, CM_wreg_col).Value = replace(FSET_code, "_", "")
				ObjExcel.Cells(excel_row, CM_abawd_col).Value = replace(ABAWD_code, "_", "")

                Call BULK_ABAWD_FSET_exemption_finder
                If snap_status = "INACTIVE" then report_notes = report_notes & "Don't assign - Inactive. "
            End if
        End if
    End if
    ObjExcel.Cells(excel_row, notes_col).Value = report_notes
    excel_row = excel_row + 1
    PMI_number = ""
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

FOR i = 1 to 15		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review the TLR list.")
