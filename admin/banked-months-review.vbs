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
    objExcel.DisplayAlerts = alerts_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url,,,, my_password) 'Opens an excel file from a specific URL
    ''(file.Path,,,, "mypassword",,,,,,,,,,)

end function

function sort_dates(dates_array)

    dim ordered_dates ()
    redim ordered_dates(0)

    days =  0
    do

        prev_date = ""
        for each thing in dates_array
            check_this_date = TRUE
            For each known_date in ordered_dates
                if known_date = thing Then check_this_date = FALSE
                'MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
            next
            if check_this_date = TRUE Then
                if prev_date = "" Then
                    prev_date = thing
                Else
                    if DateDiff("d", prev_date, thing) <0 then
                        prev_date = thing
                    end if
                end if
            end if
        next
        if prev_date <> "" Then
            redim preserve ordered_dates(days)
            ordered_dates(days) = prev_date
            days = days + 1
        end if
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

Function review_ABAWD_FSET_exemptions(person_ref_nbr, possible_exemption, exemption_array)
    '--- This function screens for ABAWD/FSET exemptions for SNAP.
    '===== Keywords: MAXIS, ABAWD, FSET, exemption, SNAP
    CALL check_for_MAXIS(False)

    possible_exemption = FALSE
    exemption_list = ""

    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	IF person_ref_nbr <> "" THEN
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen cl_age, 2, 8, 76
		IF cl_age = "  " THEN cl_age = 0
		cl_age = cl_age * 1
		IF cl_age < 18 OR cl_age >= 50 THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have exemption. Age = " & cl_age & "."
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
					exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have disability exemption. DISA end date = " & disa_end_dt & "."
					disa_status = True
				END IF
			ELSE
				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
					exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have disability exemption. DISA has no end date."
					disa_status = True
				END IF
			END IF
			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
				IF DateDiff("D", date, cert_end_dt) > 0 THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have disability exemption. DISA Certification end date = " & cert_end_dt & "."
			ELSE
				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
					EMReadScreen cert_begin_dt, 8, 7, 47
					IF cert_begin_dt <> "__ __ __" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have disability exemption. DISA certification has no end date."
				END IF
			END IF
		END IF
	END IF

    '>>>>>>>>>>>> EATS GROUP
	CALL navigate_to_MAXIS_screen("STAT", "EATS")
	eats_group_members = ""
	memb_found = True
	EMReadScreen all_eat_together, 1, 4, 72
	IF all_eat_together = "_" THEN
		eats_group_members = "01" & ","
        memb_found = FALSE
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
		'IF placeholder_HH_array <> eats_group_members THEN script_end_procedure("You are asking the script to verify ABAWD and SNAP E&T exemptions for a household that does not match the EATS group. The script cannot support this request. It will now end." & "&~&" & "Please re-run the script selecting only the individuals in the EATS group.")
		eats_group_members = trim(eats_group_members)
		eats_group_members = split(eats_group_members, ",")

		IF all_eat_together <> "_" THEN
			CALL write_value_and_transmit("MEMB", 20, 71)
			FOR EACH eats_pers IN eats_group_members
				IF eats_pers <> "" AND person_ref_nbr <> eats_pers THEN
					CALL write_value_and_transmit(eats_pers, 20, 76)
					EMReadScreen cl_age, 2, 8, 76
					IF cl_age = "  " THEN cl_age = 0
						cl_age = cl_age * 1
						IF cl_age =< 17 THEN
							exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": May have exemption for minor child caretaker. Household member " & eats_pers & " is minor. Please review for accuracy."
						END IF
				END IF
			NEXT
		END IF

		CALL write_value_and_transmit("DISA", 20, 71)
		FOR EACH disa_pers IN eats_group_members
			disa_status = false
			IF disa_pers <> "" AND disa_pers <> person_ref_nbr THEN
				CALL write_value_and_transmit(disa_pers, 20, 76)
				EMReadScreen num_of_DISA, 1, 2, 78
				IF num_of_DISA <> "0" THEN
					EMReadScreen disa_end_dt, 10, 6, 69
					disa_end_dt = replace(disa_end_dt, " ", "/")
					EMReadScreen cert_end_dt, 10, 7, 69
					cert_end_dt = replace(cert_end_dt, " ", "/")
					IF IsDate(disa_end_dt) = True THEN
						IF DateDiff("D", date, disa_end_dt) > 0 THEN
							exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": MAY have an exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
							disa_status = TRUE
						END IF
					ELSEIF IsDate(disa_end_dt) = False THEN
						IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
    							exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & " : MAY have exemption for disabled household member. Member " & disa_pers & " DISA end date = " & disa_end_dt & "."
							disa_status = true
						END IF
					END IF
					IF IsDate(cert_end_dt) = True AND disa_status = False THEN
						IF DateDiff("D", date, cert_end_dt) > 0 THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification end date = " & cert_end_dt & "."
					ELSE
						IF (cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999") THEN
							EMReadScreen cert_begin_dt, 8, 7, 47
							IF cert_begin_dt <> "__ __ __" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person & ": MAY have exemption for disabled household member. Member " & disa_pers & " DISA certification has no end date."
						END IF
					END IF
				END IF
			END IF
		NEXT
	END IF

    '>>>>>>>>>>>>>>EARNED INCOME
	IF person_ref_nbr <> "" THEN
		prosp_inc = 0
		prosp_hrs = 0
		prospective_hours = 0

		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
        'MsgBox "At JOBS"
		EMWritescreen person_ref_nbr, 20, 76
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
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
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
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen num_of_RBIC, 1, 2, 78
		IF num_of_RBIC <> "0" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Has RBIC panel. Please review for ABAWD and/or SNAP E&T exemption."
		IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
			exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to be working 30 hours/wk (regardless of wage level) or earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions."
		ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
			exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption and SNAP E&T exemptions."
		END IF
	END IF

    '>>>>>>>>>>>>UNEA
    CALL navigate_to_MAXIS_screen("STAT", "UNEA")
	IF person_ref_nbr <> "" THEN
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen num_of_UNEA, 1, 2, 78
		IF num_of_UNEA <> "0" THEN
			DO
				EMReadScreen unea_type, 2, 5, 37
				EMReadScreen unea_end_dt, 8, 7, 68
				unea_end_dt = replace(unea_end_dt, " ", "/")
				IF IsDate(unea_end_dt) = True THEN
					IF DateDiff("D", date, unea_end_dt) > 0 THEN
						IF unea_type = "14" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
					END IF
				ELSE
					IF unea_end_dt = "__/__/__" THEN
						IF unea_type = "14" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions."
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		END IF
	END IF

    '>>>>>>>>>PBEN
    CALL navigate_to_MAXIS_screen("STAT", "PBEN")
	IF person_ref_nbr <> "" THEN
		EMWriteScreen "PBEN", 20, 71
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen num_of_PBEN, 1, 2, 78
		IF num_of_PBEN <> "0" THEN
			pben_row = 8
			DO
			    IF pben_type = "12" THEN		'UI pending'
					EMReadScreen pben_disp, 1, pben_row, 77
					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
						exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption."
						EXIT DO
					END IF
				ELSE
					pben_row = pben_row + 1
				END IF
			LOOP UNTIL pben_row = 14
		END IF
	END IF

    '>>>>>>>>>>PREG
    CALL navigate_to_MAXIS_screen("STAT", "PREG")
	IF person_ref_nbr <> "" THEN
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen num_of_PREG, 1, 2, 78
        EMReadScreen preg_due_dt, 8, 10, 53
        preg_due_dt = replace(preg_due_dt, " ", "/")
		EMReadScreen preg_end_dt, 8, 12, 53

		IF num_of_PREG <> "0" THen
            If preg_due_dt <> "__/__/__" Then
                If DateDiff("d", date, preg_due_dt) > 0 AND preg_end_dt = "__ __ __" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have active pregnancy. Please review for ABAWD exemption."
                If DateDiff("d", date, preg_due_dt) < 0 Then exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have an overdue pregnancy, person may meet a minor child exemption. Contact client."
            End If
        End If
    END IF

    '>>>>>>>>>>PROG
    CALL navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash1_status, 4, 6, 74
    EMReadScreen cash2_status, 4, 7, 74
    IF cash1_status = "ACTV" OR cash2_status = "ACTV" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " Case is active on CASH programs. Please review for ABAWD and SNAP E&T exemption."

    '>>>>>>>>>>ADDR
    CALL navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen homeless_code, 1, 10, 43
    EmReadscreen addr_line_01, 16, 6, 43

    IF homeless_code = "Y" or addr_line_01 = "GENERAL DELIVERY" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " Client is claiming homelessness. If client has barriers to employment, they could meet the 'Unfit for Employment' exemption. Exemption began 05/2018."

    '>>>>>>>>>SCHL/STIN/STEC
	IF person_ref_nbr <> "" THEN
        CALL navigate_to_MAXIS_screen("STAT", "SCHL")
		CALL write_value_and_transmit(person_ref_nbr, 20, 76)
		EMReadScreen num_of_SCHL, 1, 2, 78
		IF num_of_SCHL = "1" THEN
			EMReadScreen school_status, 1, 6, 40
			IF school_status <> "N" THEN exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions."
		ELSE
			EMWriteScreen "STIN", 20, 71
			CALL write_value_and_transmit(person_ref_nbr, 20, 76)
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
							exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have active student income. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
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
				CALL write_value_and_transmit(person_ref_nbr, 20, 76)
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
								exemption_list = exemption_list & "&~&" & "* " & MAXIS_footer_month & "/" & MAXIS_footer_year & " M" & person_ref_nbr & ": Appears to have active student expenses. Please review student status to confirm SNAP eligibility as well as ABAWD and SNAP E&T exemptions."
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

    IF exemption_list = "" THEN
    	'exemption_list = "*** NOTICE!!! ***" & "&~&" & "It appears there are NO missed exemptions for ABAWD or SNAP E&T in MAXIS for this case. The script has checked ADDR, EATS, MEMB, DISA, JOBS, BUSI, RBIC, UNEA, PREG, PROG, PBEN, SCHL, STIN, and STEC for member(s) " & household_persons & "." & "&~&" & "Please make sure you are carefully reviewing the client's case file for any exemption-supporting documents."
    ELSE
        possible_exemption = TRUE
    	exemption_list = "The script has checked for ABAWD and SNAP E&T exemptions coded in MAXIS for MEMB " & person_ref_nbr & "." & exemption_list
        exemption_array = split(exemption_list, "&~&")
    END IF

    'Displaying the results...now with added MsgBox bling.
    'vbSystemModal will keep the results in the foreground.
    'MsgBox exemption_list, vbInformation + vbSystemModal, "ABAWD/FSET Exemption Check -- Results"
End Function

function find_three_ABAWD_months(all_counted_months)

    Call navigate_to_MAXIS_screen("STAT", "WREG")
    Call write_value_and_transmit(HH_memb, 20, 76)

    'Opening the Excel file
    Set objABAWDExcel = CreateObject("Excel.Application")
    objABAWDExcel.Visible = True
    Set objWorkbook = objABAWDExcel.Workbooks.Add()
    objABAWDExcel.DisplayAlerts = True

    'Changes name of Excel sheet to the case number
    objABAWDExcel.ActiveSheet.Name = "#" & MAXIS_case_number

    'adding column header information to the Excel list
    objABAWDExcel.Cells(1, 1).Value = "Month"
    objABAWDExcel.Cells(1, 2).Value = "MEMB " & HH_memb
    objABAWDExcel.Cells(1, 3).Value = "SNAP"
    objABAWDExcel.Cells(1, 4).Value = "GA"
    objABAWDExcel.Cells(1, 5).Value = "MFIP"
    objABAWDExcel.Cells(1, 6).Value = "MF - FS"
    objABAWDExcel.Cells(1, 7).Value = "DWP"
    objABAWDExcel.Cells(1, 8).Value = "RCA"
    objABAWDExcel.Cells(1, 9).Value = "MSA"

    'formatting the cells
    'FOR i = 1 to col_to_use
    FOR i = 1 to 9
        objABAWDExcel.Cells(1, i).Font.Bold = True		'bold font
        objABAWDExcel.Columns(i).AutoFit()				'sizing the columns
        objABAWDExcel.columns(i).NumberFormat = "@" 		'formatting as text
    NEXT

    excel_row = 2



    EmWriteScreen "x", 13, 57		'Pulls up the WREG tracker'
    transmit
    EMREADScreen tracking_record_check, 15, 4, 40  		'adds cases to the rejection list if the ABAWD tracking record cannot be accessed.
    If tracking_record_check <> "Tracking Record" then abawd_gather_error = abawd_gather_error & vbNewLine & "Unable to enter ABAWD tracking record of member " & HH_memb
    bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
    bene_yr_row = 10
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

        'counted date year: this is found on rows 7-11. Row 11 is current year plus one, so this will be exclude this list.
        If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", date), 2)
        If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, date)), 2)
        If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, date)), 2)
        If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)

        EMReadScreen counted_date_year, 2, bene_yr_row, 14								'reading counted year date
        abawd_counted_months_string = counted_date_month & "/" & counted_date_year		'creating new date variable

        objABAWDExcel.Cells(excel_row, 1).Value = abawd_counted_months_string

        'reading to see if a month is counted month or not
        EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
        IF is_counted_month <> "_" then objABAWDExcel.Cells(excel_row, 2).Value = is_counted_month
        If is_counted_month = "X" OR is_counted_month = "M" Then
            If counted_month_one = "" Then
                counted_month_one = abawd_counted_months_string
            ElseIf counted_month_two = "" Then
                counted_month_two = abawd_counted_months_string
            ElseIf counted_month_three = "" Then
                counted_month_three = abawd_counted_months_string
            End If
        End If
        excel_row = excel_row + 1

        bene_mo_col = bene_mo_col - 4		're-establishing serach once the end of the row is reached
        IF bene_mo_col = 15 THEN
            bene_yr_row = bene_yr_row - 1
            bene_mo_col = 63
        END IF
    LOOP until bene_yr_row = 6

    PF3 	'to exit the ABAWD tracking record

    '--------------------------------------------------------------------------------------------------------------------------------------------------INQX
    INQX_yr = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)

    Call navigate_to_MAXIS_screen("MONY", "INQX")
    EMWritescreen "01", 6, 38
    EMWritescreen INQX_yr, 6, 41
    EMWritescreen CM_mo, 6, 53
    EMwritescreen CM_yr, 6, 56
    EMWritescreen "X", 9, 5		'Snap
    EMWritescreen "X", 10, 5	'MFIP
    EMWritescreen "X", 11, 5 	'GA
    EMWritescreen "X", 15, 5	'RCA
    EMWritescreen "X", 13, 50	'MSA
    EMWritescreen "X", 17, 50 	'DWP
    transmit

    EMReadScreen no_issuance, 11, 24, 2
    If no_issuance = "NO ISSUANCE" then abawd_gather_error = abawd_gather_error & vbNewLine & HH_memb & " does not have any issuance during this period. The script will now end."
    one_page = FALSE        'Reset for the loop

    EMReadScreen single_page, 8, 17, 73
    If trim(single_page) = "" then
        one_page = True
    Else
        PF8
        EMReadScreen single_page_again, 8, 17, 73
        If trim(single_page) = trim(single_page_again) then one_page = True
    End if

    'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
    Do
        PF7
        EMReadScreen first_page_check, 20, 24, 2
    LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page

    Excel_row = 2
    DO
        row = 6				'establishing the row to start searching for issuance
        tracking_month = objABAWDExcel.cells(excel_row, 1).Value	're-establishing the case number to use for the case
        If trim(tracking_month) = "" then exit do

        Do
            Do
                EMReadScreen issuance_month, 2, row, 73
                EMReadScreen issuance_year, 2, row, 79
                EMReadScreen issuance_day, 2, row, 65
                INQX_issuance = issuance_month & "/" & issuance_year
                If trim(INQX_issuance) = "" then exit do

                If tracking_month = INQX_issuance then
                    EMReadScreen prog_type, 5, row, 16
                    prog_type = trim(prog_type)
                    EMReadScreen amt_issued, 7, row, 40
                    If issuance_day <> "01" then amt_issued = amt_issued & "*"
                    If prog_type = "FS" 	then fs_issued = fs_issued + amt_issued
                    If prog_type = "GA" 	then ga_issued = ga_issued + amt_issued
                    If prog_type = "MF-MF" 	then mfip_issued = mfip_issued + amt_issued
                    If prog_type = "MF-FS" 	then mffs_issued = mffs_issued + amt_issued
                    If prog_type = "DW" 	then dw_issued = dw_issued + amt_issued
                    If prog_type = "RC" 	then rc_issued = rc_issued + amt_issued
                    If prog_type = "MS" 	then ms_issued = ms_issued + amt_issued
                End if
                row = row + 1
            Loop until row = 18

            If one_page = True then exit do
            PF8
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check = "CAN NOT PAGE THROUGH " then
                review_required = True
                last_page = True
            elseIf last_page_check = "THIS IS THE LAST PAGE" then
                last_page = True
            Else
                last_page = False
                row = 6		're-establishes row for the new page
            End if
        Loop until last_page = True

        objABAWDExcel.Cells(excel_row, 3).Value = fs_issued
        objABAWDExcel.Cells(excel_row, 4).Value = ga_issued
        objABAWDExcel.Cells(excel_row, 5).Value = mfip_issued
        objABAWDExcel.Cells(excel_row, 6).Value = mffs_issued
        objABAWDExcel.Cells(excel_row, 7).Value = dw_issued
        objABAWDExcel.Cells(excel_row, 8).Value = rc_issued
        objABAWDExcel.Cells(excel_row, 9).Value = ms_issued

        amt_issued = ""
        fs_issued = ""
        ga_issued = ""
        mfip_issued = ""
        mffs_issued = ""
        dw_issued = ""
        rc_issued = ""
        ms_issued = ""

        If one_page <> True then
            'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
            Do
                PF7
                EMReadScreen first_page_check, 20, 24, 2
            LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page
        End if

        excel_row = excel_row + 1
    Loop

    FOR i = 1 to 9
        objABAWDExcel.Columns(i).AutoFit()				'sizing the columns
    NEXT

    BeginDialog Dialog1, 0, 0, 141, 90, "Confirm Counted ABAWD Months"
      EditBox 30, 30, 30, 15, counted_month_one
      EditBox 30, 50, 30, 15, counted_month_two
      EditBox 30, 70, 30, 15, counted_month_three
      ButtonGroup ButtonPressed
        OkButton 85, 70, 50, 15
      Text 10, 5, 135, 20, "The script has determined that the counted ABAWD months appear to be:"
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1

        Loop until err_msg = ""
        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    If counted_month_one = "" OR counted_month_two = "" OR counted_month_three = "" Then
        ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
    End If

    all_counted_months = ""
    If counted_month_one <> "" Then all_counted_months = counted_month_one
    If counted_month_two <> "" Then
        If all_counted_months = "" THen
            all_counted_months = counted_month_two
        Else
            all_counted_months = all_counted_months & "~" & counted_month_two
        End If
    End If
    If counted_month_three <> "" Then
        If all_counted_months = "" THen
            all_counted_months = counted_month_three
        Else
            all_counted_months = all_counted_months & "~" & counted_month_three
        End If
    End If

    ObjExcel.Cells(list_row, counted_ABAWD_col).Value = all_counted_months

    objABAWDExcel.DisplayAlerts = FALSE
    objABAWDExcel.Quit
    objABAWDExcel.DisplayAlerts = TRUE
    Set objABAWDExcel = Nothing

end function
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
'Const BM_to_approve_col = 17
Const counted_ABAWD_col = 16
Const NOT_BANKED_col    = 17

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

Const mo_one_type       = 15        'PLUS 9
Const mo_two_type       = 16
Const mo_three_type     = 17
Const mo_four_type      = 18
Const mo_five_type      = 19
Const mo_six_type       = 20
Const mo_seven_type     = 21
Const mo_eight_type     = 22
Const mo_nine_type      = 23

Const mo_one_update     = 24        'PLUS 18
Const mo_two_update     = 25
Const mo_three_update   = 26
Const mo_four_update    = 27
Const mo_five_update    = 28
Const mo_six_update     = 29
Const mo_seven_update   = 30
Const mo_eight_update   = 31
Const mo_nine_update    = 32

Const mo_one_app        = 33        'PLUS 27
Const mo_two_app        = 34
Const mo_three_app      = 35
Const mo_four_app       = 36
Const mo_five_app       = 37
Const mo_six_app        = 38
Const mo_seven_app      = 39
Const mo_eight_app      = 40
Const mo_nine_app       = 41

Const clt_curr_mo_stat  = 42
Const case_errors       = 43
Const used_ABAWD_mos    = 44
Const cm_approval_type  = 45    'This month the type of ABAWD/exclusion
Const nm_approval_type  = 46    'The next month the type of ABAWD/exclusion
Const remove_case       = 47
Const months_to_approve = 48

'TYPES OF SNAP/ABAWD months
  ' "INACTIVE"
  ' "EXEMPT"
  ' "BANKED MONTH"
  ' "REG ABAWD"
  ' "PRORATED"

'==========================================================================================================================


'THE SCRIPT================================================================================================================

'Connects to BlueZone
EMConnect ""

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)

Dim BANKED_MONTHS_CASES_ARRAY ()
ReDim BANKED_MONTHS_CASES_ARRAY (months_to_approve, 0)

Dim CASE_ABAWD_TO_COUNT_ARRAY ()
ReDim CASE_ABAWD_TO_COUNT_ARRAY (months_to_approve, 0)

CALL back_to_SELF
EmReadscreen MX_region, 10, 22, 48
MX_region = trim(MX_region)

'Initial Dialog will have worker select which option is going to be run at this time
    'Assess Banked Month cases from DAIL PEPR List
    'Review monthly BOBI report of all SNAP clients
    'Review of Banked Months cases
    'Approve ongoing Banked Month Cases
    'HAVE DEVELOPER MODE

'IF NOT in Developer Mode, check to be sure we are in production

BeginDialog Dialog1, 0, 0, 181, 80, "Dialog"
  'DropListBox 15, 35, 160, 45, "Find ABAWD Months", process_option
  DropListBox 15, 35, 160, 45, "Ongoing Banked Months Cases"+chr(9)+"Find ABAWD Months", process_option
  ButtonGroup ButtonPressed
    OkButton 70, 60, 50, 15
    CancelButton 125, 60, 50, 15
  Text 10, 10, 170, 10, "Script to assess and review Banked Months cases."
EndDialog

Do
    dialog Dialog1
    cancel_confirmation
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If process_option = "Ongoing Banked Months Cases" Then working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Ongoing banked months list.xlsx"     'THIS IS THE REAL ONE
' If process_option = "Ongoing Banked Months Cases" Then working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Copy of Ongoing banked months list.xlsx"  'use for tesing.'
'Live
If process_option = "Find ABAWD Months" Then working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Ongoing banked months list.xlsx"     'THIS IS THE REAL ONE
'Test
'If process_option = "Find ABAWD Months" Then working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Banked Months\Ongoing banked months list.xlsx"     'THIS IS THE REAL ONE

If process_option = "Ongoing Banked Months Cases" Then
    'For REAL'
    working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Ongoing banked months list.xlsx"     'THIS IS THE REAL ONE
    ' working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Copy of Ongoing banked months list.xlsx"  'use for tesing.'
    ' 'For testing'
    ' working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Banked Months\Ongoing banked months list.xlsx"

    'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
    call excel_open_pw(working_excel_file_path, True, False, ObjExcel, objWorkbook, "BM")

    ' 'TRYING TO FIGURE OUT HOW TO IDENTIFY IF EXCEL IS IN READ ONLY AND IF TO OPEN OR NOT'
    ' Set ObjExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
    ' ObjExcel.Visible = True
    ' ObjExcel.DisplayAlerts = True
    ' Set objWorkbook = ObjExcel.Workbooks.Open(working_excel_file_path,,,, "BM",,,,,,FALSE) 'Opens an excel file from a specific URL

    'Open( FileName , UpdateLinks , ReadOnly , Format , Password , WriteResPassword , IgnoreReadOnlyRecommended , Origin , Delimiter , Editable , Notify , Converter , AddToMru , Local , CorruptLoad )

    'ObjExcel.SendKeys "{ENTER}"
    'ObjExcel.SendKeys "{RETURN}"
    'ObjExcel.SendKeys "~"

    ObjExcel.Worksheets("Ongoing banked months").Activate
ElseIf process_option = "Find ABAWD Months" Then
    working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Ongoing banked months list.xlsx"     'THIS IS THE REAL ONE
    call excel_open(working_excel_file_path, True, False, ObjExcel, objWorkbook)
    ObjExcel.Worksheets("Ongoing banked months").Activate
End If

If process_option = "Ongoing Banked Months Cases" Then
    excel_row_to_start = objExcel.Cells(1, 20).Value
    excel_row_to_start = excel_row_to_start * 1
    excel_row_to_start = excel_row_to_start + 1
    excel_row_to_start = excel_row_to_start & ""
Else
    excel_row_to_start = "2"
    'stop_time = "1"
End If

BeginDialog Dialog1, 0, 0, 176, 140, "Dialog"
  EditBox 25, 55, 30, 15, stop_time
  EditBox 65, 100, 30, 15, excel_row_to_start
  EditBox 65, 120, 30, 15, excel_row_to_end
  ButtonGroup ButtonPressed
    OkButton 115, 120, 50, 15
  Text 5, 10, 165, 10, "This run of the script will review and help process: "
  Text 5, 20, 165, 10, process_option
  Text 10, 35, 140, 20, "To time limit the run of the script enter the numeber of hours to run the script:"
  Text 65, 60, 50, 10, "Hours"
  Text 10, 80, 145, 20, "The run can be limited by indicating which rows of the Excel file to review/process:"
  Text 15, 105, 50, 10, "Excel to start"
  Text 15, 125, 45, 10, "Excel to end"
EndDialog

Do
    Do
        err_msg = ""
        dialog Dialog1

        If trim(stop_time) <> "" AND IsNumeric(stop_time) = FALSE Then err_msg = err_msg & vbNewLine & "- Number of hours should be a number."
        If trim(excel_row_to_start) <> "" AND IsNumeric(excel_row_to_start) = FALSE Then err_msg = err_msg & vbNewLine & "- Start row of Excel should be a number."
        If trim(excel_row_to_end) <> "" AND IsNumeric(excel_row_to_end) = FALSE Then err_msg = err_msg & vbNewLine & "- End row of Excel should be a number."

        If err_msg <> "" Then MsgBox "** Please Resolve the Following to Continue:" & vbNew & err_msg

    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If trim(excel_row_to_start) = "" Then excel_row_to_start = 2
excel_row_to_start = excel_row_to_start * 1
If trim(excel_row_to_end) <> "" Then excel_row_to_end = excel_row_to_end * 1

'making stop time a number
If stop_time <> "" Then
    stop_time = FormatNumber(stop_time, 2,          0,                 0,                      0)
                            'number     dec places  leading 0 - FALSE    neg nbr in () - FALSE   use deliminator(comma) - FALSE
    stop_time = stop_time * 60 * 60     'tunring hours to seconds

    end_time = timer + stop_time        'timer is the number of seconds from 12:00 AM so we need to add the hours to run to the time to determine at what point the script should exit the loop
Else
    end_time = 84600    'sets the end time for 11:30 PM so that is doesn't end out
End If

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

If process_option = "Find ABAWD Months" Then
    list_row = excel_row_to_start           'script will allow the user to set where the script will start in taking case information from the excel row
    the_case = 0                            'setting the incrementer for adding to the array
    Do
        If trim(ObjExcel.Cells(list_row, counted_ABAWD_col).Value) = "" Then
            ReDim Preserve CASE_ABAWD_TO_COUNT_ARRAY(months_to_approve, the_case)
            CASE_ABAWD_TO_COUNT_ARRAY(case_nbr, the_case)           = trim(ObjExcel.Cells(list_row, case_nbr_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_excel_row, the_case)      = list_row
            CASE_ABAWD_TO_COUNT_ARRAY(memb_ref_nbr, the_case)       = trim(ObjExcel.Cells(list_row, memb_nrb_col).Value)

            CASE_ABAWD_TO_COUNT_ARRAY(clt_last_name, the_case)      = trim(ObjExcel.Cells(list_row, last_name_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_first_name, the_case)     = trim(ObjExcel.Cells(list_row, first_name_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)          = trim(ObjExcel.Cells(list_row, notes_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_one, the_case)         = trim(ObjExcel.Cells(list_row, first_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_two, the_case)         = trim(ObjExcel.Cells(list_row, scnd_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_three, the_case)       = trim(ObjExcel.Cells(list_row, third_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_four, the_case)        = trim(ObjExcel.Cells(list_row, fourth_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_five, the_case)        = trim(ObjExcel.Cells(list_row, fifth_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_six, the_case)         = trim(ObjExcel.Cells(list_row, sixth_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_svn, the_case)         = trim(ObjExcel.Cells(list_row, svnth_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_eight, the_case)       = trim(ObjExcel.Cells(list_row, eighth_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_mo_nine, the_case)        = trim(ObjExcel.Cells(list_row, ninth_mo_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(clt_curr_mo_stat, the_case)   = trim(ObjExcel.Cells(list_row, curr_mo_stat_col).Value)
            CASE_ABAWD_TO_COUNT_ARRAY(months_to_approve, the_case)  = ""    'set this to zero at every run as it should be handled prior to the script run

            If excel_row_to_end = list_row Then Exit DO

            list_row = list_row + 1     'incrementing the excel row and the array
            the_case = the_case + 1
        Else
            If excel_row_to_end = list_row Then Exit DO
            list_row = list_row + 1
        End If
    Loop Until trim(ObjExcel.Cells(list_row, case_nbr_col).Value) = ""  'end of the list has case number as blank

    'Loop through each item in the array to review the case.
    For the_case = 0 to UBOUND(CASE_ABAWD_TO_COUNT_ARRAY, 2)
        MAXIS_case_number = CASE_ABAWD_TO_COUNT_ARRAY(case_nbr, the_case)
        HH_memb = CASE_ABAWD_TO_COUNT_ARRAY(memb_ref_nbr, the_case)
        list_row = CASE_ABAWD_TO_COUNT_ARRAY(clt_excel_row, the_case)

        counted_month_one = ""
        counted_month_two = ""
        counted_month_three = ""

        Updates_made = FALSE

        continue_search = TRUE
        'establishing what MAXIS_footer_month and year are for WREG panel/ATR months determination
        MAXIS_footer_month 	= CM_mo
        MAXIS_footer_year 	= CM_yr

        Call navigate_to_MAXIS_screen("STAT", "WREG")

        'Checking for PRIV cases.
        EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
        IF priv_check = "PRIVIL" THEN
            CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "PRIV " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
            ObjExcel.Cells(list_row, notes_col).Value = CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
            continue_search = FALSE
        ELSE
            Call write_value_and_transmit(HH_memb, 20, 76)

            EMReadScreen wreg_total, 1, 2, 78
            If wreg_total = "0" then
                CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "NO WREG " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                ObjExcel.Cells(list_row, notes_col).Value = CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                continue_search = FALSE
            End If
        END IF


        If continue_search = TRUE THen

            abawd_gather_error = ""
            Call find_three_ABAWD_months(counted_list)          'made all of the below into a function because we need it in another process as well
            If abawd_gather_error <> "" Then
                MsgBox "Review this case as script could not gather Information to assist in ABAWD months determination." & vbNewLine & abawd_gather_error
                CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "FIND ABAWD MONTHS MANUALLY " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
            End If
            ' 'Opening the Excel file
            ' Set objABAWDExcel = CreateObject("Excel.Application")
            ' objABAWDExcel.Visible = True
            ' Set objWorkbook = objABAWDExcel.Workbooks.Add()
            ' objABAWDExcel.DisplayAlerts = True
            '
            ' 'Changes name of Excel sheet to the case number
            ' objABAWDExcel.ActiveSheet.Name = "#" & MAXIS_case_number
            '
            ' 'adding column header information to the Excel list
            ' objABAWDExcel.Cells(1, 1).Value = "Month"
            ' objABAWDExcel.Cells(1, 2).Value = "MEMB " & HH_memb
            ' objABAWDExcel.Cells(1, 3).Value = "SNAP"
            ' objABAWDExcel.Cells(1, 4).Value = "GA"
            ' objABAWDExcel.Cells(1, 5).Value = "MFIP"
            ' objABAWDExcel.Cells(1, 6).Value = "MF - FS"
            ' objABAWDExcel.Cells(1, 7).Value = "DWP"
            ' objABAWDExcel.Cells(1, 8).Value = "RCA"
            ' objABAWDExcel.Cells(1, 9).Value = "MSA"
            '
            ' 'formatting the cells
            ' 'FOR i = 1 to col_to_use
            ' FOR i = 1 to 9
            ' 	objABAWDExcel.Cells(1, i).Font.Bold = True		'bold font
            ' 	objABAWDExcel.Columns(i).AutoFit()				'sizing the columns
            ' 	objABAWDExcel.columns(i).NumberFormat = "@" 		'formatting as text
            ' NEXT
            '
            ' excel_row = 2
            '
            ' 'establishing what MAXIS_footer_month and year are for WREG panel/ATR months determination
            ' MAXIS_footer_month 	= CM_mo
            ' MAXIS_footer_year 	= CM_yr
            '
            ' EmWriteScreen "x", 13, 57		'Pulls up the WREG tracker'
            ' transmit
            ' EMREADScreen tracking_record_check, 15, 4, 40  		'adds cases to the rejection list if the ABAWD tracking record cannot be accessed.
            ' If tracking_record_check <> "Tracking Record" then script_end_procedure("Unable to enter ABAWD tracking record of member.")
            ' bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
            ' bene_yr_row = 10
            ' DO
            '     'establishing variables for specific ABAWD counted month dates
            '     If bene_mo_col = "19" then counted_date_month = "01"
            '     If bene_mo_col = "23" then counted_date_month = "02"
            '     If bene_mo_col = "27" then counted_date_month = "03"
            '     If bene_mo_col = "31" then counted_date_month = "04"
            '     If bene_mo_col = "35" then counted_date_month = "05"
            '     If bene_mo_col = "39" then counted_date_month = "06"
            '     If bene_mo_col = "43" then counted_date_month = "07"
            '     If bene_mo_col = "47" then counted_date_month = "08"
            '     If bene_mo_col = "51" then counted_date_month = "09"
            '     If bene_mo_col = "55" then counted_date_month = "10"
            '     If bene_mo_col = "59" then counted_date_month = "11"
            '     If bene_mo_col = "63" then counted_date_month = "12"
            '
            '     'counted date year: this is found on rows 7-11. Row 11 is current year plus one, so this will be exclude this list.
            '     If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", date), 2)
            '     If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, date)), 2)
            '     If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, date)), 2)
            '     If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)
            '
            ' 	EMReadScreen counted_date_year, 2, bene_yr_row, 14								'reading counted year date
            ' 	abawd_counted_months_string = counted_date_month & "/" & counted_date_year		'creating new date variable
            '
            ' 	objABAWDExcel.Cells(excel_row, 1).Value = abawd_counted_months_string
            '
            ' 	'reading to see if a month is counted month or not
            ' 	EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
            ' 	IF is_counted_month <> "_" then objABAWDExcel.Cells(excel_row, 2).Value = is_counted_month
            '     If is_counted_month = "X" OR is_counted_month = "M" Then
            '         If counted_month_one = "" Then
            '             counted_month_one = abawd_counted_months_string
            '         ElseIf counted_month_two = "" Then
            '             counted_month_two = abawd_counted_months_string
            '         ElseIf counted_month_three = "" Then
            '             counted_month_three = abawd_counted_months_string
            '         End If
            '     End If
            ' 	excel_row = excel_row + 1
            '
            ' 	bene_mo_col = bene_mo_col - 4		're-establishing serach once the end of the row is reached
            ' 	IF bene_mo_col = 15 THEN
            ' 		bene_yr_row = bene_yr_row - 1
            ' 		bene_mo_col = 63
            ' 	END IF
            ' LOOP until bene_yr_row = 6
            '
            ' PF3 	'to exit the ABAWD tracking record
            '
            ' '--------------------------------------------------------------------------------------------------------------------------------------------------INQX
            ' INQX_yr = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)
            '
            ' Call navigate_to_MAXIS_screen("MONY", "INQX")
            ' EMWritescreen "01", 6, 38
            ' EMWritescreen INQX_yr, 6, 41
            ' EMWritescreen CM_mo, 6, 53
            ' EMwritescreen CM_yr, 6, 56
            ' EMWritescreen "X", 9, 5		'Snap
            ' EMWritescreen "X", 10, 5	'MFIP
            ' EMWritescreen "X", 11, 5 	'GA
            ' EMWritescreen "X", 15, 5	'RCA
            ' EMWritescreen "X", 13, 50	'MSA
            ' EMWritescreen "X", 17, 50 	'DWP
            ' transmit
            '
            ' EMReadScreen no_issuance, 11, 24, 2
            ' If no_issuance = "NO ISSUANCE" then script_end_procedure(HH_memb & " does not have any issuance during this period. The script will now end.")
            ' one_page = FALSE        'Reset for the loop
            '
            ' EMReadScreen single_page, 8, 17, 73
            ' If trim(single_page) = "" then
            ' 	one_page = True
            ' Else
            ' 	PF8
            ' 	EMReadScreen single_page_again, 8, 17, 73
            ' 	If trim(single_page) = trim(single_page_again) then one_page = True
            ' End if
            '
            ' 'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
            ' Do
            ' 	PF7
            ' 	EMReadScreen first_page_check, 20, 24, 2
            ' LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page
            '
            ' Excel_row = 2
            ' DO
            ' 	row = 6				'establishing the row to start searching for issuance
            ' 	tracking_month = objABAWDExcel.cells(excel_row, 1).Value	're-establishing the case number to use for the case
            ' 	If trim(tracking_month) = "" then exit do
            '
            ' 	Do
            ' 	    Do
            ' 	    	EMReadScreen issuance_month, 2, row, 73
            ' 	    	EMReadScreen issuance_year, 2, row, 79
            ' 			EMReadScreen issuance_day, 2, row, 65
            ' 	    	INQX_issuance = issuance_month & "/" & issuance_year
            ' 	    	If trim(INQX_issuance) = "" then exit do
            '
            ' 	    	If tracking_month = INQX_issuance then
            ' 	    		EMReadScreen prog_type, 5, row, 16
            ' 	    		prog_type = trim(prog_type)
            ' 	    		EMReadScreen amt_issued, 7, row, 40
            ' 				If issuance_day <> "01" then amt_issued = amt_issued & "*"
            ' 	    		If prog_type = "FS" 	then fs_issued = fs_issued + amt_issued
            ' 	    		If prog_type = "GA" 	then ga_issued = ga_issued + amt_issued
            ' 	    		If prog_type = "MF-MF" 	then mfip_issued = mfip_issued + amt_issued
            ' 	    		If prog_type = "MF-FS" 	then mffs_issued = mffs_issued + amt_issued
            ' 	    		If prog_type = "DW" 	then dw_issued = dw_issued + amt_issued
            ' 	    		If prog_type = "RC" 	then rc_issued = rc_issued + amt_issued
            ' 	    		If prog_type = "MS" 	then ms_issued = ms_issued + amt_issued
            ' 	    	End if
            ' 	    	row = row + 1
            ' 	    Loop until row = 18
            '
            ' 		If one_page = True then exit do
            ' 		PF8
            ' 		EMReadScreen last_page_check, 21, 24, 2
            ' 		If last_page_check = "CAN NOT PAGE THROUGH " then
            ' 		 	review_required = True
            ' 			last_page = True
            ' 		elseIf last_page_check = "THIS IS THE LAST PAGE" then
            ' 			last_page = True
            ' 		Else
            ' 			last_page = False
            ' 			row = 6		're-establishes row for the new page
            ' 		End if
            ' 	Loop until last_page = True
            '
            ' 	objABAWDExcel.Cells(excel_row, 3).Value = fs_issued
            ' 	objABAWDExcel.Cells(excel_row, 4).Value = ga_issued
            ' 	objABAWDExcel.Cells(excel_row, 5).Value = mfip_issued
            ' 	objABAWDExcel.Cells(excel_row, 6).Value = mffs_issued
            ' 	objABAWDExcel.Cells(excel_row, 7).Value = dw_issued
            ' 	objABAWDExcel.Cells(excel_row, 8).Value = rc_issued
            ' 	objABAWDExcel.Cells(excel_row, 9).Value = ms_issued
            '
            ' 	amt_issued = ""
            ' 	fs_issued = ""
            ' 	ga_issued = ""
            ' 	mfip_issued = ""
            ' 	mffs_issued = ""
            ' 	dw_issued = ""
            ' 	rc_issued = ""
            ' 	ms_issued = ""
            '
            ' 	If one_page <> True then
            ' 	    'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
            ' 	    Do
            ' 	    	PF7
            ' 	    	EMReadScreen first_page_check, 20, 24, 2
            ' 	    LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page
            ' 	End if
            '
            ' 	excel_row = excel_row + 1
            ' Loop
            '
            ' FOR i = 1 to 9
            ' 	objABAWDExcel.Columns(i).AutoFit()				'sizing the columns
            ' NEXT
            '
            ' BeginDialog Dialog1, 0, 0, 141, 90, "Confirm Counted ABAWD Months"
            '   EditBox 30, 30, 30, 15, counted_month_one
            '   EditBox 30, 50, 30, 15, counted_month_two
            '   EditBox 30, 70, 30, 15, counted_month_three
            '   ButtonGroup ButtonPressed
            '     OkButton 85, 70, 50, 15
            '   Text 10, 5, 135, 20, "The script has determined that the counted ABAWD months appear to be:"
            ' EndDialog
            '
            ' Do
            '     Do
            '         err_msg = ""
            '
            '         dialog Dialog1
            '
            '     Loop until err_msg = ""
            '     call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
            ' LOOP UNTIL are_we_passworded_out = false
            '
            ' If counted_month_one = "" OR counted_month_two = "" OR counted_month_three = "" Then
            '     ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
            ' End If
            '
            ' all_counted_months = ""
            ' If counted_month_one <> "" Then all_counted_months = counted_month_one
            ' If counted_month_two <> "" Then
            '     If all_counted_months = "" THen
            '         all_counted_months = counted_month_two
            '     Else
            '         all_counted_months = all_counted_months & "~" & counted_month_two
            '     End If
            ' End If
            ' If counted_month_three <> "" Then
            '     If all_counted_months = "" THen
            '         all_counted_months = counted_month_three
            '     Else
            '         all_counted_months = all_counted_months & "~" & counted_month_three
            '     End If
            ' End If
            '
            ' ObjExcel.Cells(list_row, counted_ABAWD_col).Value = all_counted_months
            '
            ' objABAWDExcel.DisplayAlerts = FALSE
            ' objABAWDExcel.Quit
            ' objABAWDExcel.DisplayAlerts = TRUE
            ' Set objABAWDExcel = Nothing

        END IF
        'This will cause the script to end if there was a timer set and the script needs to end
        If timer > end_time Then
            end_msg = "Success! Script has run for " & stop_time/60/60 & " hours and has finished." & vbNewLine & "The script processed the rows " & CASE_ABAWD_TO_COUNT_ARRAY(clt_excel_row, 0) & " through " & CASE_ABAWD_TO_COUNT_ARRAY(clt_excel_row, Ubound(CASE_ABAWD_TO_COUNT_ARRAY, 2))
            Exit For
        Else
            end_msg = "Script run completed. The script processed the rows "  & CASE_ABAWD_TO_COUNT_ARRAY(clt_excel_row, 0) & " through " & CASE_ABAWD_TO_COUNT_ARRAY(clt_excel_row, Ubound(CASE_ABAWD_TO_COUNT_ARRAY, 2))
        End If

    Next
End If

'This is to handle cases that were already approved as a banked month and needs to be continually reviewed and approved every month
If process_option = "Ongoing Banked Months Cases" Then
    Dim ABAWD_MONTHS_ARRAY
    list_row = excel_row_to_start           'script will allow the user to set where the script will start in taking case information from the excel row
    the_case = 0                            'setting the incrementer for adding to the array
    Do
        ReDim Preserve BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case)
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
        BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)     = trim(ObjExcel.Cells(list_row, counted_ABAWD_col).Value)
        BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case)  = ""    'set this to zero at every run as it should be handled prior to the script run

        list_row = list_row + 1     'incrementing the excel row and the array
        the_case = the_case + 1

        If excel_row_to_end = list_row Then Exit DO

    Loop Until trim(ObjExcel.Cells(list_row, case_nbr_col).Value) = ""  'end of the list has case number as blank

    'Loop through each item in the array to review the case.
    For the_case = 0 to UBOUND(BANKED_MONTHS_CASES_ARRAY, 2)
        ABAWD_MONTHS_ARRAY = ""
        still_three_used = TRUE
        BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = FALSE
        MAXIS_footer_month = ""
        MAXIS_footer_year = ""
        other_notes = ""

        list_row = BANKED_MONTHS_CASES_ARRAY(clt_excel_row, the_case)       'setting the excel row to what was found in the array
        MAXIS_case_number = BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case)   'setting the case number to this variable for nav functions to work
        BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) = Right("00"&BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 2)    'formatting the member number to be 2 digit
        ObjExcel.Cells(list_row, memb_nrb_col) = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)                              'adding the formatted number to the excel sheet because I am tired of crazy looking excel files
        HH_memb = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)
        'list_of_exemption = ""
        start_month = ""    'blanking out these variables for each loop through the array
        start_year = ""
        assist_a_new_approval = FALSE
        case_note_done = FALSE

        ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 6

        ' If ObjExcel.Cells(list_row, NOT_BANKED_col).Value = "TRUE" Then
        '     MAXIS_footer_month = CM_mo
        '     MAXIS_footer_year = CM_yr
        '
        '     Call back_to_SELF
        '
        '     Call navigate_to_MAXIS_screen
        ' End If

        ' For month_indicator = clt_mo_one to clt_mo_nine     'These are set as constants that are numbers (parameters in the array) so we can loop through them
        month_indicator = clt_mo_one
        Do
            Call back_to_SELF                               'need to go to SELF so we can go to a different month
            month_tracked = FALSE                           'this is reset for every month
            abawd_status = ""                               'blanking out each variable
            this_month_is_ABAWD = FALSE
            fset_wreg_status = ""
            approvable_month = FALSE
            BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = ""

            If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) <> "" Then                          'if the spreadsheet already has a month listed in one of the 'tracked months'
                MAXIS_footer_month = left(BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case), 2)      'we set the footer month and year  using the month from the spreadsheet so that we look at the right month in STAT
                MAXIS_footer_year = right(BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case), 2)

                month_tracked = TRUE            'if the month was listed on the spreadsheet - it was already tracked
            Else
                If month_indicator = clt_mo_one Then
                    If MAXIS_footer_month = CM_mo AND MAXIS_footer_year = CM_yr Then
                        MAXIS_footer_month = CM_plus_1_mo
                        MAXIS_footer_year = CM_plus_1_yr
                    Else
                        MAXIS_footer_month = CM_mo
                        MAXIS_footer_year = CM_yr
                    End If
                Else
                    first_of_footer_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year     'there was no month in the spreadsheet
                    next_month = DateAdd("m", 1, first_of_footer_month)                         'the month is advanded by ONE from what the last month we looked at was

                    MAXIS_footer_month = DatePart("m", next_month)          'formatting the month and year and setting them for the nav functions to work
                    MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)

                    MAXIS_footer_year = DatePart("yyyy", next_month)
                    MAXIS_footer_year = right(MAXIS_footer_year, 2)
                End If
            End If

            If BANKED_MONTHS_CASES_ARRAY(clt_curr_mo_stat, the_case) <> "" Then
                last_updated_mo = left(BANKED_MONTHS_CASES_ARRAY(clt_curr_mo_stat, the_case), 5)
                last_updated_mo = left(last_updated_mo, 2) & "/01/" & right(last_updated_mo, 2)

                Do
                    the_now_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year

                    If DateDiff("m", the_now_month, last_updated_mo) >= 0 Then
                        If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = "" Then

                            first_of_footer_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year     'there was no month in the spreadsheet
                            next_month = DateAdd("m", 1, first_of_footer_month)                         'the month is advanded by ONE from what the last month we looked at was

                            MAXIS_footer_month = DatePart("m", next_month)          'formatting the month and year and setting them for the nav functions to work
                            MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)

                            MAXIS_footer_year = DatePart("yyyy", next_month)
                            MAXIS_footer_year = right(MAXIS_footer_year, 2)

                        Else
                            month_indicator = month_indicator + 1

                            If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) <> "" Then                          'if the spreadsheet already has a month listed in one of the 'tracked months'
                                MAXIS_footer_month = left(BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case), 2)      'we set the footer month and year  using the month from the spreadsheet so that we look at the right month in STAT
                                MAXIS_footer_year = right(BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case), 2)

                                month_tracked = TRUE            'if the month was listed on the spreadsheet - it was already tracked
                            Else
                                first_of_footer_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year     'there was no month in the spreadsheet
                                next_month = DateAdd("m", 1, first_of_footer_month)                         'the month is advanded by ONE from what the last month we looked at was

                                MAXIS_footer_month = DatePart("m", next_month)          'formatting the month and year and setting them for the nav functions to work
                                MAXIS_footer_month = right("00"&MAXIS_footer_month, 2)

                                MAXIS_footer_year = DatePart("yyyy", next_month)
                                MAXIS_footer_year = right(MAXIS_footer_year, 2)
                            End If
                        End If
                    End If
                    the_now_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
                Loop until DateDiff("m", the_now_month, last_updated_mo) < 0
                If MAXIS_footer_month = CM_plus_2_mo AND MAXIS_footer_year = CM_plus_2_yr Then Exit DO
            End If
            ' MsgBox "Footer Month - " & MAXIS_footer_month & "/" & MAXIS_footer_year

            If MAXIS_footer_month = CM_mo AND MAXIS_footer_year = CM_yr Then approvable_month = TRUE
            If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then approvable_month = TRUE

            Do
                Call navigate_to_MAXIS_screen("STAT", "SUMM")
                EmReadscreen summ_check, 4, 2, 46
            Loop until summ_check = "SUMM"

            client_not_in_HH = FALSE
            If HH_memb <> "01" Then
                Call navigate_to_MAXIS_screen("STAT", "REMO")
                Call write_value_and_transmit(HH_memb, 20, 76)
                EmReadscreen check_if_memb_in_HH, 33, 24, 2
                If check_if_memb_in_HH = "MEMBER " & HH_memb & " IS NOT IN THE HOUSEHOLD" Then
                    client_not_in_HH = TRUE
                Else
                    EmReadscreen HH_memb_left_date, 8, 8, 53
                    EmReadscreen HH_memb_exp_return, 8, 14, 53
                    EmReadscreen HH_memb_actual_return, 8, 16, 53

                    If HH_memb_left_date <> "__ __ __" and HH_memb_exp_return = "__ __ __" and HH_memb_actual_return = "__ __ __" Then client_not_in_HH = TRUE
                End If
            End If

            If client_not_in_HH = TRUE Then
                Call navigate_to_MAXIS_screen("STAT", "MEMB")
                EMSetCursor 4, 33
                PF1
                memb_list_row = 9
                Do
                    EmReadscreen the_ref_number, 2, memb_list_row, 5
                    If the_ref_number = HH_memb Then
                        EmReadscreen clt_pmi, 8, memb_list_row, 49
                        clt_pmi = trim(clt_pmi)
                        Exit Do
                    End If

                    memb_list_row = memb_list_row + 1
                    If memb_list_row = 19 Then
                        PF8
                        memb_list_row = 9
                    End If
                Loop until the_ref_number = "  "

                Call back_to_SELF
                Call Navigate_to_MAXIS_screen("PERS", "    ")

                EmWriteScreen clt_pmi, 15, 36
                transmit

                'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 6

                BeginDialog Dialog1, 0, 0, 216, 135, "Dialog"
                  CheckBox 10, 60, 180, 10, "Check here if client is active SNAP on another case.", new_case_checkbox
                  EditBox 95, 75, 65, 15, new_case_number
                  CheckBox 10, 100, 180, 10, "Check here if client is not active SNAP on any case.", clt_closed_checkbox
                  ButtonGroup ButtonPressed
                    OkButton 160, 115, 50, 15
                  Text 10, 10, 205, 10, "It appears that MEMBER 00 has been removed from this case."
                  Text 10, 25, 185, 25, "The script has navigated to a person search for the PMI associated with this member. Review client status and review if client is active SNAP on another case."
                  Text 20, 80, 65, 10, "New case number:"
                EndDialog

                Do
                    Do
                        err_msg = ""

                        dialog Dialog1

                        new_case_number = trim(new_case_number)
                        If new_case_checkbox = checked AND clt_closed_checkbox = checked Then err_msg = err_msg & vbNewLine & "* Client cannot be both on a new case and completely inactive SNAP. Select one box to check."
                        If new_case_checkbox = unchecked AND clt_closed_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Indicate if client is still active on SNAP or not. Select one of the boxes."
                        If new_case_checkbox = checked AND new_case_number = "" Then err_msg = err_msg & vbNewLine & "* Since the client is active on a new case, the new case number needs to be entered here."

                        If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

                    Loop until err_msg = ""
                    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
                LOOP UNTIL are_we_passworded_out = false

                'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 0

                If clt_closed_checkbox = checked Then
                    BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = TRUE
                    month_indicator = month_indicator + 1
                    BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "INACTIVE"   'Type of ABAWD/SNAP month
                    BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = FALSE       'WREG to be updated
                    BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE       'SNAP approval to be made

                    If InStr(BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case), " CLIENT WAS REMOVED FROM THIS CASE AND IS NO LONGER ACTIVE SNAP.") = 0 Then BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " CLIENT WAS REMOVED FROM THIS CASE AND IS NO LONGER ACTIVE SNAP."
                End If
                If new_case_checkbox = checked Then
                    BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " Client moved from SNAP on " & MAXIS_case_number & " to SNAP on case number " & new_case_number & " in the month " & MAXIS_footer_month & "/" & MAXIS_footer_year & "."
                    MAXIS_case_number = new_case_number
                    BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case) = MAXIS_case_number
                    client_not_in_HH = FALSE

                    list_row = BANKED_MONTHS_CASES_ARRAY(clt_excel_row, the_case)       'setting the excel row to what was found in the array
                    ObjExcel.Cells(list_row, case_nbr_col) = BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case)                              'adding the formatted number to the excel sheet because I am tired of crazy looking excel files

                    Call back_to_SELF
                    Call navigate_to_MAXIS_screen("STAT", "MEMB")

                    EMSetCursor 4, 33
                    PF1
                    memb_list_row = 9
                    Do
                        EmReadscreen the_pmi, 8, memb_list_row, 49
                        the_pmi = trim(the_pmi)
                        If the_pmi = clt_pmi Then
                            EmReadscreen ref_nbr, 2, memb_list_row, 5
                            Exit Do
                        End If

                        memb_list_row = memb_list_row + 1
                        If memb_list_row = 19 Then
                            PF8
                            memb_list_row = 9
                        End If
                    Loop until the_pmi = "  "

                    BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) = ref_nbr                                                             'resetting this reference number to the one on the new case
                    HH_memb = ref_nbr
                    ObjExcel.Cells(list_row, memb_nrb_col) = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)                              'adding the formatted number to the excel sheet because I am tired of crazy looking excel files

                End If
            End If

            If client_not_in_HH = FALSE Then

                'TODO add functionality to note on the spreadhseet and in the case for cases in which the used banked month has been CONFRIMED - that the client was active in a past or current month.
                ' MsgBox BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)
                If BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = "" Then
                    abawd_gather_error = ""
                    Call find_three_ABAWD_months(counted_list)
                    If abawd_gather_error <> "" Then
                        MsgBox "Review this case as script could not gather Information to assist in ABAWD months determination." & vbNewLine & abawd_gather_error
                        CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "FIND ABAWD MONTHS MANUALLY " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                        'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
                    End If
                    BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = counted_list
                End If
                If BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) <> "" THen
                    If InStr(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), "~") <> 0 Then
                        ABAWD_MONTHS_ARRAY = Split(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), "~")
                    End If

                    If Ubound(ABAWD_MONTHS_ARRAY) <> 2 Then
                        BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = ""
                        abawd_gather_error = ""
                        Call find_three_ABAWD_months(counted_list)
                        If abawd_gather_error <> "" Then
                            MsgBox "Review this case as script could not gather Information to assist in ABAWD months determination." & vbNewLine & abawd_gather_error
                            CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "FIND ABAWD MONTHS MANUALLY " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                            'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
                        End If
                        BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = counted_list

                        ABAWD_MONTHS_ARRAY = ""

                        If BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = "" Then
                            CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "PROCESS MANUALLY " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                            'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
                            Exit Do
                        Else
                            If InStr(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), "~") <> 0 Then
                                ABAWD_MONTHS_ARRAY = Split(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), "~")
                            Else
                                CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "PROCESS MANUALLY " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                                'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
                                Exit Do
                            End If

                            ' MsgBox "UBOUND - " & UBOUND(ABAWD_MONTHS_ARRAY)
                            If Ubound(ABAWD_MONTHS_ARRAY) <> 2 Then
                                CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case) = "PROCESS MANUALLY " & CASE_ABAWD_TO_COUNT_ARRAY(clt_notes, the_case)
                                'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
                                Exit Do
                            End If
                        End If

                    End If

                    For each used_month in ABAWD_MONTHS_ARRAY
                        the_month = left(used_month, 2)
                        the_year = right(used_month, 2)
                        the_ABAWD_month = the_month & "/01/" & the_year

                        used_month = the_ABAWD_month
                    Next

                    Call sort_dates(ABAWD_MONTHS_ARRAY)

                    For each used_month in ABAWD_MONTHS_ARRAY
                        the_month = right("00"&DatePart("m", used_month), 2)
                        the_year = right(DatePart("yyyy", used_month), 2)

                        used_month = the_month & "/" & the_year
                    Next
                    still_three_used = TRUE

                    BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = ""
                    For each used_month in ABAWD_MONTHS_ARRAY
                        ' MsgBox used_month
                        the_month = left(used_month, 2)
                        the_year = right(used_month, 2)
                        the_ABAWD_month = the_month & "/01/" & the_year

                        this_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
                        ' MsgBox "The ABAWD month is " & the_ABAWD_month & vbNewLine & "Difference is " & DateDiff("m", the_ABAWD_month, this_month)

                        'TODO need to address this in each month to be reviewed since we may be looking at more than one month'
                        If still_three_used = TRUE Then
                            If DateDiff("m", the_ABAWD_month, this_month) > 35 Then
                                still_three_used = FALSE
                                BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " " & the_ABAWD_month & " is more than 36 months ago."
                                BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = ""
                                BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "REG ABAWD"   'Type of ABAWD/SNAP month
                                BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE       'WREG to be updated
                                If approvable_month = FALSE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE       'SNAP approval to be made
                                If approvable_month = TRUE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = TRUE       'SNAP approval to be made

                                BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) & "~" & MAXIS_footer_month & "/" & MAXIS_footer_year
                                removed_month = the_month & "/" & the_year
                                this_month_is_ABAWD = FALSE
                            Else
                                If BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = "" Then
                                    BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = the_month & "/" & the_year
                                Else
                                    BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) & "~" & the_month & "/" & the_year
                                End If
                            End If
                        Else
                            If BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = "" Then
                                BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = the_month & "/" & the_year
                            Else
                                BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) & "~" & the_month & "/" & the_year
                            End If
                        End If
                    Next
                    If left(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), 1) = "~" Then BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = right(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), len(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case))-1)
                    ObjExcel.Cells(list_row, counted_ABAWD_col).Value = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)
                End If

                Call navigate_to_MAXIS_screen("CASE", "PERS")       'go to CASE/PERS - which is month specific
                pers_row = 10                                       'the first member number starts at row 10
                clt_SNAP_status = ""                                'blanking out this variable for each loop through the array
                Do
                    EMReadScreen pers_ref_numb, 2, pers_row, 3      'reading the member number
                    If pers_ref_numb = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) Then   'compaing it to the member number in the array
                        EMReadScreen clt_SNAP_status, 1, pers_row, 54       'if it matches then read the SNAP status
                        Exit Do
                    Else                                            'if it doesn't match
                        pers_row = pers_row + 3                     'go to the next member number - which is 3 rows down
                        If pers_row = 19 Then                       'if it reaches 19 - this is further down from the last member
                            PF8                                     'go to the next page and reset to line 10
                            pers_row = 10
                        End If
                    End If
                Loop until pers_ref_numb = "  "                     'this is the end of the list

                If clt_SNAP_status = "A" Then                       'If the member number was listed as ACTIVE on CASE/PERS then the script will check STAT
                    Call back_to_SELF
                    Call Navigate_to_MAXIS_screen ("MONY", "INQX")
                    EmWriteScreen MAXIS_footer_month, 6, 38
                    EmWriteScreen MAXIS_footer_year, 6, 41
                    EmWriteScreen MAXIS_footer_month, 6, 53
                    EmWriteScreen MAXIS_footer_year, 6, 56
                    EmWriteScreen "X", 9, 5

                    transmit

                    mony_row = 6
                    Do
                        EmReadscreen from_mo, 2, mony_row, 62
                        If from_mo = MAXIS_footer_month Then
                            EmReadscreen from_day, 2, mony_row, 65
                            Exit Do
                        End If
                        mony_row = mony_row + 1
                    Loop until from_mo = "  "


                    If from_day <> "01" AND from_day <> "  " Then
                        If BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "REG ABAWD" Then
                            BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "PRORATED - REG"
                        Else
                            BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "PRORATED - BM"
                        End If
                        ' MsgBox "Month type - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) & vbNewLine & "List of ABAWD - " & BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) & vbNewLine & "Instring calc - " & InStr(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), MAXIS_footer_month & "/" & MAXIS_footer_year)
                        If InStr(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), MAXIS_footer_month & "/" & MAXIS_footer_year) <> 0 Then
                            ABAWD_MONTHS_ARRAY = ""
                            ABAWD_MONTHS_ARRAY = Split(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), "~")

                            BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = ""
                            For each used_month in ABAWD_MONTHS_ARRAY
                                ' MsgBox used_month
                                the_month = left(used_month, 2)
                                the_year = right(used_month, 2)

                                If MAXIS_footer_month = the_month AND MAXIS_footer_year = the_year Then
                                    BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) & "~" & removed_month
                                Else
                                    BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)  = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) & "~" & the_month & "/" & the_year
                                End If
                            Next

                            If left(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), 1) = "~" Then BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case) = right(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case), len(BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case))-1)
                            ObjExcel.Cells(list_row, counted_ABAWD_col).Value = BANKED_MONTHS_CASES_ARRAY(used_ABAWD_mos, the_case)
                        End If

                        BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE
                        BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = ""
                        BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & "~ " &MAXIS_footer_month & "/" & MAXIS_footer_year & " is a PRORATED MONTH."
                    End If

                    Call navigate_to_MAXIS_screen("STAT", "WREG")   'Go to WREG - where ABAWD information is

                    If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = "" Then
                        BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE
                        If approvable_month = FALSE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE       'SNAP approval to be made
                        If approvable_month = TRUE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = TRUE       'SNAP approval to be made
                    End If
                    month_tracker_nbr = month_indicator - 5 & ""

                    EMWriteScreen BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 20, 76 'go to the panel for the correct member
                    transmit

                    EMReadScreen fset_wreg_status, 2, 8, 50     'Reading the FSET Status and ABAWD status
                    EMReadScreen abawd_status, 2, 13, 50

                    If fset_wreg_status = "30" AND abawd_status = "13" Then
                        EmReadscreen banked_counter, 1, 14, 50
                        If banked_counter <> month_tracker_nbr Then
                            BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE
                            If approvable_month = FALSE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE       'SNAP approval to be made
                            If approvable_month = TRUE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = TRUE       'SNAP approval to be made
                        End If
                    End If

                    ' MsgBox "BEGINNING" & vbNewLine & "Month type - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) & vbNewLine & "The month is - '" & BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) & "'" & vbNewLine & "Update WREG - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) & vbNewLine & "Do Approval - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case)

                    If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = "" Then
                        'MsgBox "Pause"
                        If fset_wreg_status = "30" Then
                            Call review_ABAWD_FSET_exemptions(BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), exemption_exists, list_of_exemption)

                            code_for_banked = TRUE      'resetting this variable
                            If exemption_exists = TRUE Then     'if the function above finds a potential issue then the script will ask the worker to determine if it is supposed to still be BANKED

                                'finding the height of the dialog
                                dlg_len = 130
                                For each exemption in list_of_exemption
                                    hgt = 10
                                    if len(exemption) > 100 then hgt = 20
                                    if len(exemption) > 200 then hgt = 30
                                    dlg_len = dlg_len + hgt
                                Next
                                y_pos = 75

                                'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 6

                                'This dialog will list all of the exemptions the function found
                                BeginDialog Dialog1, 0, 0, 346, dlg_len, "Possible ABAWD/FSET Exemption"
                                'BeginDialog Dialog1, 0, 0, 346, 135, "Possible ABAWD/FSET Exemption"
                                  GroupBox 15, 10, 325, 55, "Case Review"
                                  Text 60, 25, 250, 10, "*** THIS CASE NEEDS REVIEW OF POSSIBLE ABAWD EXEMPTION ***"
                                  Text 20, 40, 310, 20, "At this time, review this case as STAT indicates that the client may meet an ABAWD exemption and may no longer need to use Banked Months. Check the case and update now."
                                  For each exemption in list_of_exemption
                                    'Text 10, 75, 330, 10, "exemption list"
                                    hgt = 10
                                    if len(exemption) > 100 then hgt = 20
                                    if len(exemption) > 200 then hgt = 30
                                    Text 10, y_pos, 330, hgt, exemption
                                    y_pos = y_pos + hgt + 5
                                  next
                                  Text 70, y_pos, 205, 10, "*** IF THIS CASE MEETS AN ABAWD OR FSET EXEMPTION ***"
                                  y_pos = y_pos + 10
                                  Text 90, y_pos, 160, 10, "*** UPDATE AND DO A NEW APPROVAL NOW ***"
                                  y_pos = y_pos + 15
                                  ButtonGroup ButtonPressed
                                    PushButton 15, y_pos, 145, 15, "CASE STILL NEEDS BANKED MONTHS", still_banked_btn
                                    PushButton 165, y_pos, 165, 15, "Client now meets an ABAWD or FSET Exemption", meets_exemption_btn
                                EndDialog

                                dialog Dialog1      'display the dialog

                                'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 0

                                'If the worker indicates that the client meets an exemption this tells the script that we no longer need to code for banked months
                                If ButtonPressed = meets_exemption_btn Then
                                    BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = ""
                                    BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "EXEMPT"   'Type of ABAWD/SNAP month
                                    BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE       'WREG to be updated
                                    If approvable_month = FALSE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE       'SNAP approval to be made
                                    If approvable_month = TRUE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = TRUE       'SNAP approval to be made

                                    For each exemption in list_of_exemption
                                        BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " ~ Possible ABAWD/FSET Exemption: " & exemption
                                    Next

                                    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
                                    CALL write_value_and_transmit(BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 20, 76)
                                    EMReadScreen cl_age, 2, 8, 76
                                    IF cl_age = "  " THEN cl_age = 0
                                    cl_age = cl_age * 1
                                End If

                                If ButtonPressed = still_banked_btn Then
                                    If left(BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case), 8) = "" Then
                                    'If BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) <> "PRORATED - REG" AND BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) <> "PRORATED - BM" AND BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) <> "REG ABAWD" Then
                                        BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year
                                        BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "BANKED MONTH"   'Type of ABAWD/SNAP month
                                        BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE       'WREG to be updated
                                        If approvable_month = FALSE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE       'SNAP approval to be made
                                        If approvable_month = TRUE Then BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = TRUE       'SNAP approval to be made
                                    End If
                                End If
                            ElseIF BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "" Then
                                BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "BANKED MONTH"
                            End If
                        Else
                            BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "EXEMPT"
                            BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = FALSE
                            BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = FALSE
                            BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " ~ " & MAXIS_footer_month & "/" &  MAXIS_footer_year & " WREG coded for ABAWD exemption."
                            BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = TRUE
                        End If
                    ElseIF BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "" Then ''"PRORATED - REG" AND BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) <> "PRORATED - BM" AND BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) <> "REG ABAWD" Then
                        BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "BANKED MONTH"
                    End If

                    ' MsgBOx BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case)
                    Call back_to_SELF
                    Call navigate_to_MAXIS_screen("STAT", "WREG")   'The script or worker may have moved around in the case - need to navigate back
                    EMWriteScreen BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 20, 76
                    transmit

                    'WREG Updated'
                    EMReadScreen fset_wreg_status, 2, 8, 50     'Reading the FSET Status and ABAWD status
                    EMReadScreen abawd_status, 2, 13, 50
                    If BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE Then
                        If BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "BANKED MONTH" Then
                            If MX_region = "INQUIRY DB" Then           'If we are in Inquiry, the script runs in a developer mode, messaging the information
                                MsgBox "WREG to be updated with BM Tracker number: " & month_tracker_nbr
                                BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year
                            Else                                    'If we are in production, then we should actually update
                                PF9
                                EMWriteScreen "30", 8, 50
                                EMWriteScreen "13", 13, 50
                                EMWriteScreen month_tracker_nbr, 14, 50
                                transmit
                                EMWriteScreen "BGTX", 20, 71
                                transmit

                                BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year
                                'TODO add confirmation that WREG was updated
                            End If
                        ElseIf BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "REG ABAWD" Then
                            If MX_region = "INQUIRY DB" Then           'If we are in Inquiry, the script runs in a developer mode, messaging the information
                                MsgBox "WREG to be updated with 30/10"
                            Else                                    'If we are in production, then we should actually update
                                PF9
                                EMWriteScreen "30", 8, 50
                                EMWriteScreen "10", 13, 50
                                EMWriteScreen " ", 14, 50

                                If approvable_month = FALSE Then
                                    EMWriteScreen "X", 13, 57
                                    transmit

                                    Select Case MAXIS_footer_month
                                        Case "01"
                                            search_mo = "Jan"
                                        Case "02"
                                            search_mo = "Feb"
                                        Case "03"
                                            search_mo = "Mar"
                                        Case "04"
                                            search_mo = "Apr"
                                        Case "05"
                                            search_mo = "May"
                                        Case "06"
                                            search_mo = "Jun"
                                        Case "07"
                                            search_mo = "Jul"
                                        Case "08"
                                            search_mo = "Aug"
                                        Case "09"
                                            search_mo = "Sep"
                                        Case "10"
                                            search_mo = "Oct"
                                        Case "11"
                                            search_mo = "Nov"
                                        Case "12"
                                            search_mo = "Dec"
                                    End Select

                                    row = 1
                                    wreg_col = 1
                                    EMSearch search_mo, row, wreg_col
                                    wreg_col = wreg_col + 1

                                    wreg_row = 1
                                    col = 1
                                    EMSearch MAXIS_footer_year, wreg_row, col

                                    EMWriteScreen "M", wreg_row, wreg_col

                                    PF3
                                End If

                                transmit
                                EMWriteScreen "BGTX", 20, 71
                                transmit
                            End If
                        ElseIf BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "EXEMPT" Then
                            full_of_exemptions = JOIN(list_of_exemption, "~")
                            If InStr(full_of_exemptions, "active on CASH programs") <> 0 Then new_fset_wreg_status = "17"
                            If InStr(full_of_exemptions, "claiming homelessness") <> 0 Then new_fset_wreg_status = "03"
                            If InStr(full_of_exemptions, "minor child caretaker") <> 0 Then new_fset_wreg_status = "21"

                            If InStr(full_of_exemptions, "Age = ") <> 0 Then
                                If cl_age < 16 Then new_fset_wreg_status = "06"
                                If cl_age < 18 AND cl_age > 15 Then new_fset_wreg_status = "07"
                                If cl_age > 50 AND cl_age < 60 Then new_fset_wreg_status = "16"
                                If cl_age > 60 Then new_fset_wreg_status = "05"
                            End If
                            If InStr(full_of_exemptions, "disability exemption") <> 0 Then new_fset_wreg_status = "03"

                            If InStr(full_of_exemptions, "disabled household member") <> 0 Then new_fset_wreg_status = "04"
                            If InStr(full_of_exemptions, "Appears to be working 30 hours/wk") <> 0 Then new_fset_wreg_status = "09"

                            If InStr(full_of_exemptions, "active unemployment benefits") <> 0 Then new_fset_wreg_status = "11"
                            If InStr(full_of_exemptions, "pending, appealing, or eligible Unemployment") <> 0 Then new_fset_wreg_status = "11"


                            If InStr(full_of_exemptions, "enrolled in school") <> 0 Then new_fset_wreg_status = "12"
                            If InStr(full_of_exemptions, "active student income") <> 0 Then new_fset_wreg_status = "12"
                            If InStr(full_of_exemptions, "active student expenses") <> 0 Then new_fset_wreg_status = "12"

                            If InStr(full_of_exemptions, "active pregnancy") <> 0 Then new_abawd_status = "05"
                            If InStr(full_of_exemptions, "overdue pregnancy") <> 0 Then new_abawd_status = "05"
                            If InStr(full_of_exemptions, "Appears to be working at least 80 hours") <> 0 Then new_abawd_status = "06"

                            If new_fset_wreg_status = "21" Then new_abawd_status = "04"
                            If new_fset_wreg_status = "16" Then new_abawd_status = "03"

                            If new_fset_wreg_status = "03" OR new_fset_wreg_status = "04" OR new_fset_wreg_status = "05" OR new_fset_wreg_status = "06" OR new_fset_wreg_status = "07" OR new_fset_wreg_status = "09" OR new_fset_wreg_status = "11" OR new_fset_wreg_status = "12" Then new_abawd_status = "01"
                            If new_abawd_status = "05" OR new_abawd_status = "06" Then new_fset_wreg_status = "30"

                            'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 6

                            BeginDialog Dialog1, 0, 0, 111, 90, "FSET ABAWD Status"
                              EditBox 80, 30, 20, 15, new_fset_wreg_status
                              EditBox 80, 50, 20, 15, new_abawd_status
                              ButtonGroup ButtonPressed
                                OkButton 55, 70, 50, 15
                              Text 5, 10, 105, 20, "Confirm the FSET and ABAWD status for this client."
                              Text 5, 35, 70, 10, "FSET/WREG Status"
                              Text 5, 55, 50, 10, "ABAWD Status"
                            EndDialog

                            Do
                                err_msg = ""

                                dialog Dialog1

                                If len(new_fset_wreg_status) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter the correct FSET WREG Status."
                                If len(new_abawd_status) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter the correct ABAWD Status."

                                If err_msg <> "" Then MsgBox "** Please resolve to continue **" & vbNewLine * err_msg

                            Loop until err_msg = ""

                            'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 0


                            If MX_region = "INQUIRY DB" Then           'If we are in Inquiry, the script runs in a developer mode, messaging the information
                                MsgBox "WREG to be updated with " & new_fset_wreg_status & "/" & new_abawd_status
                            Else                                    'If we are in production, then we should actually update
                                PF9
                                EMWriteScreen new_fset_wreg_status, 8, 50
                                EMWriteScreen new_abawd_status, 13, 50
                                EMWriteScreen " ", 14, 50
                                transmit
                                EMWriteScreen "BGTX", 20, 71
                                transmit
                            End If
                        ElseIF BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "PRORATED - REG" Then
                            If MX_region = "INQUIRY DB" Then           'If we are in Inquiry, the script runs in a developer mode, messaging the information
                                MsgBox "WREG to be updated with 30/10"
                            Else                                    'If we are in production, then we should actually update
                                PF9
                                EMWriteScreen "30", 8, 50
                                EMWriteScreen "10", 13, 50
                                EMWriteScreen " ", 14, 50
                                transmit
                                EMWriteScreen "BGTX", 20, 71
                                transmit
                            End If
                        ElseIF BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "PRORATED - BM" Then
                            If MX_region = "INQUIRY DB" Then           'If we are in Inquiry, the script runs in a developer mode, messaging the information
                                MsgBox "WREG to be updated with 30/13 - tracker removed"
                            Else                                    'If we are in production, then we should actually update
                                PF9
                                EMWriteScreen "30", 8, 50
                                EMWriteScreen "13", 13, 50
                                EMWriteScreen " ", 14, 50
                                transmit
                                EMWriteScreen "BGTX", 20, 71
                                transmit
                            End If
                        End If
                    End If

                    If BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) = TRUE Then
                        If BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) = "BANKED MONTH" Then
                            other_notes = other_notes & "For the month - " & MAXIS_footer_month & "/" & MAXIS_footer_year & " the SNAP for Memb " & HH_memb & " is " & BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) & " - Banked Month: " & month_tracker_nbr & ".; "

                        Else
                            other_notes = other_notes & "For the month - " & MAXIS_footer_month & "/" & MAXIS_footer_year & " the SNAP for Memb " & HH_memb & " is " & BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) & "; "
                        End If
                        BANKED_MONTHS_CASES_ARRAY(clt_curr_mo_stat, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year & " - " & BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case)
                        Updates_made = TRUE
                    End If
                    If BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) = "INACTIVE" Then BANKED_MONTHS_CASES_ARRAY(clt_curr_mo_stat, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year & " - " & BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case)

                    If BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case) = TRUE Then
                        if start_month = "" Then
                            start_month = MAXIS_footer_month
                            start_year = MAXIS_footer_year
                        End If
                        assist_a_new_approval = TRUE
                    End If



                    '
                    ' 'Banked Months are numbered 1-9 as they are used
                    ' 'This sets the indicator for WREG using the constants from the array to determine WHICH of the month it is
                    ' month_tracker_nbr = month_indicator - 5
                    ' update_WREG = FALSE                         'resetting this boolean
                    '
                    ' If fset_wreg_status = "30" AND abawd_status = "13" Then     'If this is 30/13 then the case for this member is set as BANKED MONTHS
                    '     If MAXIS_footer_month = CM_mo AND MAXIS_footer_year = CM_yr Then update_WREG = TRUE     'if we are in the current month, we can update WREG
                    '
                    '     If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then update_WREG = TRUE   'if we are in current month plus one, we can update WREG
                    '
                    '     'In CM or CM+1 for banked months cases we will look in more detail
                    '     If update_WREG = TRUE Then
                    '         'Need to be sure that there isn't a new ABAWD month available - maybe another column with the counted months on the ongoing banked months cases
                    '         'Need to review case for possible exemption months - code from exemption finder
                    '         Call review_ABAWD_FSET_exemptions(BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), exemption_exists, list_of_exemption)
                    '
                    '         code_for_banked = TRUE      'resetting this variable
                    '         If exemption_exists = TRUE Then     'if the function above finds a potential issue then the script will ask the worker to determine if it is supposed to still be BANKED
                    '
                    '             'finding the height of the dialog
                    '             dlg_len = 130
                    '             For each exemption in list_of_exemption
                    '                 hgt = 10
                    '                 if len(exemption) > 100 then hgt = 20
                    '                 if len(exemption) > 200 then hgt = 30
                    '                 dlg_len = dlg_len + hgt
                    '             Next
                    '             y_pos = 75
                    '
                    '             ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 6
                    '
                    '             'This dialog will list all of the exemptions the function found
                    '             BeginDialog Dialog1, 0, 0, 346, dlg_len, "Possible ABAWD/FSET Exemption"
                    '             'BeginDialog Dialog1, 0, 0, 346, 135, "Possible ABAWD/FSET Exemption"
                    '               GroupBox 15, 10, 325, 55, "Case Review"
                    '               Text 60, 25, 250, 10, "*** THIS CASE NEEDS REVIEW OF POSSIBLE ABAWD EXEMPTION ***"
                    '               Text 20, 40, 310, 20, "At this time, review this case as STAT indicates that the client may meet an ABAWD exemption and may no longer need to use Banked Months. Check the case and update now."
                    '               For each exemption in list_of_exemption
                    '                 'Text 10, 75, 330, 10, "exemption list"
                    '                 hgt = 10
                    '                 if len(exemption) > 100 then hgt = 20
                    '                 if len(exemption) > 200 then hgt = 30
                    '                 Text 10, y_pos, 330, hgt, exemption
                    '                 y_pos = y_pos + hgt + 5
                    '               next
                    '               Text 70, y_pos, 205, 10, "*** IF THIS CASE MEETS AN ABAWD OR FSET EXEMPTION ***"
                    '               y_pos = y_pos + 10
                    '               Text 90, y_pos, 160, 10, "*** UPDATE AND DO A NEW APPROVAL NOW ***"
                    '               y_pos = y_pos + 15
                    '               ButtonGroup ButtonPressed
                    '                 PushButton 15, y_pos, 145, 15, "CASE STILL NEEDS BANKED MONTHS", still_banked_btn
                    '                 PushButton 165, y_pos, 165, 15, "Client now meets an ABAWD or FSET Exemption", meets_exemption_btn
                    '             EndDialog
                    '
                    '             dialog Dialog1      'display the dialog
                    '
                    '             ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 0
                    '
                    '             'If the worker indicates that the client meets an exemption this tells the script that we no longer need to code for banked months
                    '             If ButtonPressed = meets_exemption_btn Then
                    '                 code_for_banked = FALSE
                    '
                    '                 For each exemption in list_of_exemption
                    '                     BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " ~ Possible ABAWD/FSET Exemption: " & exemption
                    '                 Next
                    '             End If
                    '         End If
                    '
                    '         Call navigate_to_MAXIS_screen("STAT", "WREG")   'The script or worker may have moved around in the case - need to navigate back
                    '         EMWriteScreen BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case), 20, 76
                    '         transmit
                    '
                    '         EMReadscreen current_banked_month_indicator, 1, 14, 50      'reading what is already listed in the Banked Month indicator
                    '
                    '         'This looks to see if the BM code on WREG panel matches what we expect the tracker number to be
                    '         If right("00"&month_tracker_nbr, 1) = current_banked_month_indicator Then
                    '             code_for_banked = FALSE         'if this matches we don't need to update WREG because the tracker is already correct
                    '         Else
                    '             'This lists the months that need to be approved to be added to the Spreadsheet for manual approval
                    '             If BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case) = "" Then
                    '                 BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year
                    '             Else
                    '                 BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case) = BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case) & ", " & MAXIS_footer_month & "/" & MAXIS_footer_year
                    '             End If
                    '             'This sets the first month to be approved.
                    '             if start_month = "" Then
                    '                 start_month = MAXIS_footer_month
                    '                 start_year = MAXIS_footer_year
                    '             End If
                    '         End If
                    '         'This is for if we need the script to actually update WREG
                    '         If code_for_banked = TRUE Then
                    '             If MX_region = "INQUIRY DB" Then           'If we are in Inquiry, the script runs in a developer mode, messaging the information
                    '                 MsgBox "WREG to be updated with BM Tracker number: " & month_tracker_nbr
                    '             Else                                    'If we are in production, then we should actually update
                    '                 PF9
                    '                 EMWriteScreen month_tracker_nbr, 14, 50
                    '                 transmit
                    '                 EMWriteScreen "BGTX", 20, 71
                    '                 transmit
                    '                 'TODO add confirmation that WREG was updated
                    '             End If
                    '             assist_a_new_approval = TRUE            'This variable defines if more things should happen
                    '             'This is only reset at the beginning of the loop for each CASE - not each month
                    '         End If
                    '
                    '         'Write TIKL or something to identify cases to be approved and noted.
                    '         'IDEA - write a new column in to Excel for cases needing approval in months
                    '         'IDEA - write a process that will send a case through background and stop with a dialog to allow for manual approval.

                    'we only do an approval if we have reviewed CM + 1
                    If assist_a_new_approval = TRUE and MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then
                        Call Navigate_to_MAXIS_screen("ELIG", "FS")     'Go to ELIG in what we expect is the start month and year
                        EmWriteScreen start_month, 19, 54
                        EMWriteScreen start_year, 19, 57
                        transmit

                        'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 6
                        ' MsgBox other_notes

                        'This dialog is to assist in the noting of the approval
                        BeginDialog Dialog1, 0, 0, 236, 110, "Dialog"
                          EditBox 95, 30, 15, 15, start_month
                          EditBox 115, 30, 15, 15, start_year
                          EditBox 60, 50, 170, 15, other_notes
                          EditBox 75, 70, 155, 15, worker_signature
                          ButtonGroup ButtonPressed
                            OkButton 180, 90, 50, 15
                          Text 10, 10, 155, 20, "This case has been sent through background and ready for review and approval. "
                          Text 10, 35, 80, 10, "First Month of Approval:"
                          Text 10, 55, 45, 10, "Other Notes:"
                          Text 10, 75, 60, 10, "Worker Signature:"
                          ButtonGroup ButtonPressed
                            PushButton 10, 95, 90, 10, "No Approval Made", no_approval_button
                        EndDialog

                        dialog Dialog1

                        'ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 0

                        If ButtonPressed = OKButton Then
                            'setting the variables
                            footer_month = start_month
                            footer_year = start_year
                            Lines_in_note = ""
                            'We are going to loop through each of the months from start month to CM + 1 to gather information from ELIG
                            Do
                                Call Navigate_to_MAXIS_screen("ELIG", "SUMM")       'Go to ELIG/SUMM
                                EmWriteScreen footer_month, 19, 56                  'Go to the SNAP eligibility for the correct month and year
                                EMWriteScreen footer_year, 19, 59
                                EMWriteScreen "FS  ", 19, 71
                                transmit

                                elig_row = 7                                        'beginning of the list of members in the case
                                list_of_fs_members = ""                             'creating a list of all the members
                                Do
                                    EmReadscreen fs_memb, 2, elig_row, 10           'reading the member number, code and elig status
                                    EmReadscreen fs_memb_code, 1, elig_row, 35
                                    EmReadscreen fs_memb_elig, 8, elig_row, 57

                                    'These are when a member is active and eligible for SNAP on this case
                                    If fs_memb_code = "A" and fs_memb_elig = "ELIGIBLE" Then list_of_fs_members = list_of_fs_members & "~"& fs_memb

                                    elig_row = elig_row + 1     'looking at the next member
                                    EmReadscreen next_member, 2, elig_row, 10   'looking at if there is another member to review
                                Loop until next_member = "  "                   'This would be the end of the list of members in ELIG
                                'MsgBox "Line 947" & vbNewLine & "List of Members" & list_of_fs_members
                                If list_of_fs_members <> "" Then
                                    list_of_fs_members = right(list_of_fs_members, len(list_of_fs_members)-1)   'This was assembled from reviewing ELIG
                                    member_array = split(list_of_fs_members, "~")       'making is an ARRAY
                                End If

                                transmit    'going to FSB1'
                                transmit

                                EmReadscreen total_earned_income, 9, 8, 32
                                EmReadscreen total_unea_income, 9, 18, 32

                                total_earned_income = trim(total_earned_income)
                                total_unea_income = trim(total_unea_income)

                                If total_earned_income = "" Then total_earned_income = 0
                                If total_unea_income = "" Then total_unea_income = 0

                                total_earned_income = FormatNumber(total_earned_income, 2, -1, 0, -1)
                                total_unea_income = FormatNumber(total_unea_income, 2, -1, 0, -1)

                                transmit    'going to FSB2'

                                EmReadscreen total_shelter_costs, 9, 14, 28
                                total_shelter_costs = trim(total_shelter_costs)
                                If total_shelter_costs = "" Then total_shelter_costs = 0
                                total_shelter_costs = FormatNumber(total_shelter_costs, 2, -1, 0, -1)
                                'TODO add format number to each of these

                                transmit    'going to FSSM'

                                EmReadscreen fs_benefit_amount, 9, 13, 72
                                EmReadscreen reporting_status, 9, 8, 31

                                fs_benefit_amount = trim(fs_benefit_amount)
                                If fs_benefit_amount = "" Then fs_benefit_amount = 0
                                fs_benefit_amount = FormatNumber(fs_benefit_amount, 2, -1, 0, -1)
                                reporting_status = trim(reporting_status)
                                If fs_benefit_amount = 0 Then

                                    EmReadscreen fs_benefit_amount, 9, 10, 72
                                    fs_benefit_amount = trim(fs_benefit_amount)
                                    If fs_benefit_amount = "" Then fs_benefit_amount = 0
                                    fs_benefit_amount = FormatNumber(fs_benefit_amount, 2, -1, 0, -1)

                                End If

                                'Creating a list of each line of the case note - created here instead of adding to an array because we don't need it after the note
                                Lines_in_note = Lines_in_note & "~!~* SNAP approved for " & footer_month & "/" & footer_year
                                Lines_in_note = Lines_in_note & "~!~    Eligile Household Members: "
                                For each person in member_array
                                    Lines_in_note = Lines_in_note & person & ", "
                                Next
                                Lines_in_note = Lines_in_note & "~!~    Income: Earned: $" & total_earned_income & " Unearned: $" & total_unea_income
                                If total_shelter_costs <> "" Then  Lines_in_note = Lines_in_note & "~!~    Shelter Costs: $" & total_shelter_costs
                                Lines_in_note = Lines_in_note & "~!~    SNAP BENEFTIT: $" & fs_benefit_amount & " Reporting Status: " & reporting_status

                                first_of_footer_month = footer_month & "/01/" & footer_year     'there was no month in the spreadsheet
                                next_month = DateAdd("m", 1, first_of_footer_month)                         'the month is advanded by ONE from what the last month we looked at was

                                footer_month = DatePart("m", next_month)          'formatting the month and year and setting them for the nav functions to work
                                footer_month = right("00"&footer_month, 2)

                                footer_year = DatePart("yyyy", next_month)
                                footer_year = right(footer_year, 2)

                            Loop until footer_month = CM_plus_2_mo and footer_year = CM_plus_2_yr

                            ARRAY_OF_NOTE_LINES = split(Lines_in_note, "~!~")       'making this an array

                            case_note_done = TRUE
                            If MX_region = "INQUIRY DB" Then
                                case_note_to_display = "*** SNAP Approved for " & start_month & "/" & start_year & " ***"
                                For each note_line in ARRAY_OF_NOTE_LINES
                                    case_note_to_display = note_line
                                Next
                                case_note_to_display = case_note_to_display & vbNewLine & "* Notes: " & other_notes
                                case_note_to_display = case_note_to_display & vbNewLine & "---"
                                case_note_to_display = case_note_to_display & vbNewLine & worker_signature

                                MsgBox case_note_to_display
                            Else
                                Call start_a_blank_CASE_NOTE

                                Call write_variable_in_CASE_NOTE("*** SNAP Approved starting in " & start_month & "/" & start_year & " ***")
                                For each note_line in ARRAY_OF_NOTE_LINES
                                    Call write_variable_in_CASE_NOTE(note_line)
                                Next
                                Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
                                Call write_variable_in_CASE_NOTE("---")
                                Call write_variable_in_CASE_NOTE(worker_signature)
                            End If

                        ElseIf MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then
                            If Updates_made = TRUE Then
                                case_note_done = TRUE
                                If MX_region = "INQUIRY DB" Then
                                    case_note_to_display = "WREG Updated for ABAWD Information for M" & HH_memb
                                    notes_array = Split(other_notes, "; ")
                                    for each cnote in notes_array
                                        case_note_to_display = case_note_to_display & vbNewLine & cnote
                                    next
                                    case_note_to_display = case_note_to_display & vbNewLine & "---"
                                    case_note_to_display = case_note_to_display & vbNewLine & worker_signature

                                    MsgBox case_note_to_display
                                Else
                                    Call start_a_blank_CASE_NOTE

                                    Call write_variable_in_CASE_NOTE("WREG Updated for ABAWD Information for M" & HH_memb)
                                    Call write_bullet_and_variable_in_CASE_NOTE("Detail", other_notes)
                                    Call write_variable_in_CASE_NOTE("---")
                                    Call write_variable_in_CASE_NOTE(worker_signature)
                                End If
                            End If
                        End If
                    End If
                    '         'BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " ~ Approve SNAP for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "~"
                    '     End If
                    ' End If

                    ' If month_tracked  = FALSE Then  'if the month was not already in the traking cells from the spreadsheet
                    '     'We will add it to the array and later to the spreadsheet
                    '     BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = MAXIS_footer_month & "/" & MAXIS_footer_year
                    ' End If
                    BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = FALSE

                Else            'These cases are where the member is NOT active SNAP in the specified month
                    'If the month was tracked on the Excel spreadsheet
                    If month_tracked = TRUE Then

                        'This dialog will allow the worker to determine if this should not be tracked as a banked month '
                        BeginDialog Dialog1, 0, 0, 191, 110, "Dialog"
                          ButtonGroup ButtonPressed
                            PushButton 15, 75, 160, 15, "Yes - remove the month from the Master List", yes_remove_month_btn
                            PushButton 15, 95, 160, 10, "No - keep the month - case will be updated", no_keep_btn
                          Text 30, 10, 130, 15, "It appears that for the month " & MAXIS_footer_month & "/" & MAXIS_footer_year & " the Member " & HH_memb & " was not active on SNAP."
                          Text 30, 35, 130, 15, "This month has been tracked on the Banked Month master list."
                          Text 10, 60, 180, 10, "Should the month be removed from the tracking sheet?"
                        EndDialog

                        dialog Dialog1

                        'If the worker indicates that this should no longer be considered a USED banked month, and the array variable is removed - to be later updated in the spreadsheet
                        If ButtonPressed = yes_remove_month_btn Then BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = ""

                    Else    'if the month was not tracked - and the client not active anymore
                        BANKED_MONTHS_CASES_ARRAY(case_errors, the_case) = "STALE"      'This indicates that the case is no longer needing to be tracked
                        BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & "  ~Client is not active SNAP in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " ~  "    'adding information to NOTES on the spreadsheet
                        'MsgBox "STALE"
                    End If
                    BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = TRUE
                    BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "INACTIVE"   'Type of ABAWD/SNAP month
                End If
            End If

            IF MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr AND Updates_made = TRUE AND case_note_done = FALSE Then
                If MX_region = "INQUIRY DB" Then
                    case_note_to_display = "WREG Updated for ABAWD Information for M" & HH_memb
                    notes_array = Split(other_notes, "; ")
                    for each cnote in notes_array
                        case_note_to_display = case_note_to_display & vbNewLine & cnote
                    next
                    case_note_to_display = case_note_to_display & vbNewLine & "---"
                    case_note_to_display = case_note_to_display & vbNewLine & worker_signature

                    MsgBox case_note_to_display
                Else
                    Call start_a_blank_CASE_NOTE

                    Call write_variable_in_CASE_NOTE("WREG Updated for ABAWD Information for M" & HH_memb)
                    Call write_bullet_and_variable_in_CASE_NOTE("Detail", other_notes)
                    Call write_variable_in_CASE_NOTE("---")
                    Call write_variable_in_CASE_NOTE(worker_signature)
                End If
            End If

            ObjExcel.Cells(list_row, month_indicator).Value        = BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case)
            ' MsgBox "END" & vbNewLine & "Month type - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) & vbNewLine & "The month is - '" & BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) & "'" & vbNewLine & "Update WREG - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 18, the_case) & vbNewLine & "Do Approval - " & BANKED_MONTHS_CASES_ARRAY(month_indicator + 27, the_case)

            'MsgBox "Column " & ObjExcel.Cells(1, month_indicator) & " for tracking says - " & BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) & vbNewLine & "For the month of " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for the case: " & MAXIS_case_number & vbNewLine & "Member " & BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) & " is " & clt_SNAP_status & "." & vbNewLine & "WREG is FSET - " & fset_wreg_status & " | ABAWD - " & abawd_status
            ' If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then Exit For 'If we have completed review of CM+1, we can't gp any further and we leave the loop of all the months
            If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then Exit Do 'If we have completed review of CM+1, we can't gp any further and we leave the loop of all the months

            'NEEDS TESTING'
            If BANKED_MONTHS_CASES_ARRAY(month_indicator, the_case) = "" Then
                If month_indicator + 1 <> mo_one_type Then
                    If BANKED_MONTHS_CASES_ARRAY(month_indicator + 1, the_case) <> "" Then
                        For each_col = month_indicator to clt_mo_eight
                            BANKED_MONTHS_CASES_ARRAY(each_col, the_case) = BANKED_MONTHS_CASES_ARRAY(each_col + 1, the_case)
                            ObjExcel.Cells(list_row, each_col).Value = BANKED_MONTHS_CASES_ARRAY(each_col, the_case)
                        Next
                    End If
                End If
            End If

            If BANKED_MONTHS_CASES_ARRAY(month_indicator + 9, the_case) = "BANKED MONTH" Then month_indicator = month_indicator + 1

        ' Next
        Loop until month_indicator = clt_mo_nine

        '************************************************************************************'
        ' banked_months_tracked = TRUE
        ' function set_lastest_banked_month(date_variable, month_mo, month_yr, boo_var)
        '     month_mo = left(date_variable, 2)
        '     month_yr = right(date_variable, 2)
        '     If month_mo <> CM_plus_1_mo Then boo_var = FALSE
        '     If month_yr <> CM_plus_1_yr Then boo_var = FALSE
        ' end function
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

        'For each element in the array for the case - we are going to add that to the Excel Spreadsheet

        'MsgBox "The NOTES field will now read: " & BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case)
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
        'ObjExcel.Cells(list_row, BM_to_approve_col).Value   = BANKED_MONTHS_CASES_ARRAY(months_to_approve, the_case)

        If BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = TRUE Then
            ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 16
            ObjExcel.Cells(list_row, NOT_BANKED_col).Value = "TRUE"
        ElseIf InStr(BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case), "PROCESS MANUALLY") <> 0 Then
            ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 3
        Else
            ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex = 0
            ' ObjExcel.Range(ObjExcel.Cells(list_row, 1), ObjExcel.Cells(list_row, 17)).Interior.ColorIndex
            ' ObjExcel.Rows(list_row).Interior.ColorIndex
        End If

        objExcel.Cells(1, 20).Value = list_row

        'This will cause the script to end if there was a timer set and the script needs to end
        If timer > end_time Then
            end_msg = "Success! Script has run for " & stop_time/60/60 & " hours and has finished for the time being."
            Exit For
        Else
            end_msg = "Last case was " & MAXIS_case_number      'this is reset for testing so that I know where the script ends
        End If
    Next

End If

'NEED another spreadsheet for all cases that WERE banked months cases but are no longer - so that we can save the case information

script_end_procedure(end_msg)
