'Required for statistical purposes==========================================================================================
name_of_script = "ABAWD FSET Exemption function Update.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
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

MAXIS_case_number = "02278225"
Call test_ABAWD_FSET_exemption_finder(False, memb_number_for_bulk, eats_HH_count, snap_status, meets_childcare_exemption, homeless_exemption, best_wreg_code, best_abawd_code, verified_wreg, possible_exemptions)





function test_ABAWD_FSET_exemption_finder(bulk_run, memb_number_for_bulk, eats_HH_count, snap_status, member_age_for_bulk_run, meets_childcare_exemption, homeless_exemption, best_wreg_code, best_abawd_code, verified_wreg, possible_exemptions)
    snap_status = ""
    best_wreg_code = ""
    best_abawd_code = ""
    eats_HH_count = ""
    verified_wreg = ""
    possible_exemptions = ""
    member_age_for_bulk_run = ""

    Dim eats_group_array()
    ReDim eats_group_array(verified_abawd_const,0)

    'constants for array
    const memb_name_const           = 0
    const memb_number_const         = 1
    const memb_age_const            = 2
    const verified_exemption_const  = 3
    const potential_exempt_const    = 4
    const verified_wreg_const       = 5
    const verified_abawd_const      = 6

    entry_record = 0

    call TLR_determine_SNAP_unit(eats_group_members, memb_found, eats_HH_count)

    'sets up array for the exemption and potential exemption checks.
    eats_group_members = trim(eats_group_members)
    eats_group_members = split(eats_group_members, ",")

    If bulk_run Then
        eats_group_array(memb_number_const, entry_record) = memb_number_for_bulk
    Else
        For each memb in eats_group_members
            If trim(memb) <> "" then
                ReDim Preserve eats_group_array(verified_abawd_const, entry_record)	'This resizes the array based on the number of members
                eats_group_array(memb_number_const, entry_record) = memb
                entry_record = entry_record + 1			'This increments to the next entry in the array'
                stats_counter = stats_counter + 1
            End if
        Next
    End If

    Call TLR_active_progs_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, snap_status)

    Call TLR_demographic_exemptions(eats_group_members, eats_group_array, memb_number_const, memb_name_const, verified_exemption_const, verified_wreg_const, potential_exempt_const, meets_childcare_exemption)
    If bulk_run Then member_age_for_bulk_run = eats_group_array(memb_age_const, 0)

    Call TLR_disability_exemptions(eats_group_members, eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const)

    Call TLR_employed_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, verified_abawd_const)

    Call TLR_unearned_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, potential_exempt_const)

    Call TLR_pending_benefit_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, potential_exempt_const)

    Call TLR_preg_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const)

    Call TLR_homeless_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, potential_exempt_const)

    Call TLR_school_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const)

    If bulk_run Then
		age_50_thru_59 = False
		age_50_thru_55 = False 'This boolean is used for case note workaround. If anyone 50-55, we note that in case note the date of TLR is not specific to 11/01/25 implementation date.
		age_60_thru_64 = False


		If eats_group_array(memb_age_const, 0) => 50 then
			If eats_group_array(memb_age_const, 0) =< 59 then age_50_thru_59 = True
			If eats_group_array(memb_age_const, 0) =< 55 then age_50_thru_55 = True
		End if
		'----------------------------------------------------------------------------------------------------'05 - Age 60 or older
        If eats_group_array(memb_age_const, 0) => 60 then
			If eats_group_array(memb_age_const, 0) < 65 then
                age_60_thru_64  = True
			End if
		End if


	    'filter the list here for best_wreg_code
	    If trim(eats_group_array(verified_wreg_const, 0)) = "" then
	    	best_wreg_code = "30"
            If verified_abawd = "" then
	    		best_abawd_code = "10"
	    	Else
	    		best_abawd_code = verified_abawd 'this should only be 06 for now but maybe more later
	    	End if
	    Elseif len(eats_group_array(verified_wreg_const, 0)) = 3 then
            best_wreg_code = left(eats_group_array(verified_wreg_const, 0),2) 'resetting variable
        Else
			'Multiple wreg heirarchies possible based on the person situations and system workarounds required.
			If age_50_thru_59 = True then
				wreg_hierarchy = array("03","04","05","06","07","08","09","10","11","12","13","14","20","15","21","17","23","30","16")
			ElseIf age_60_thru_64 = True then
				wreg_hierarchy = array("03","04","06","07","08","09","10","11","12","13","14","20","15","16","21","17","23","30","05")
			Elseif meets_childcare_exemption = False then
				wreg_hierarchy = array("03","04","05","06","07","08","09","10","11","12","13","14","20","15","16","17","23","30")
			Else
				'this is for non-workarounds
				wreg_hierarchy = array("03","04","05","06","07","08","09","10","11","12","13","14","20","15","16","21","17","23","30")
			End if

            for each code in wreg_hierarchy
                If instr(verified_wreg, code) then
                    best_wreg_code = code
                    exit for
                End if
            next
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
            If best_wreg_code = "30" then best_abawd_code = "09" 'This is for native exemption folks only since that is the only thing we can read for in MAXIS to determine the verified_wreg code. Otherwise anyone who is TLR the verified_wreg is "".
        End If

        verified_wreg = eats_group_array(verified_wreg_const, 0)
        possible_exemptions = eats_group_array(potential_exempt_const, 0)
    Else
        exemption_message = ""
        For items = 0 to UBound(eats_group_array, 2)
            If trim(eats_group_array(verified_exemption_const, items)) = "" then eats_group_array(verified_exemption_const, items) = "N/A"
            If trim(eats_group_array(potential_exempt_const, items)) = "" then eats_group_array(potential_exempt_const, items) = "N/A"
            exemption_message = exemption_message & "------------------------------ " & vbcr & "MEMB #" & eats_group_array(memb_number_const, items) & eats_group_array(memb_name_const, items) & ":" & vbcr & "------------------------------ " & vbcr & _
            "Verified Exemptions: " & eats_group_array(verified_exemption_const, items) & vbcr & "Potential Exemptions: " & eats_group_array(potential_exempt_const, items) & vbcr
        Next

        MsgBox exemption_message, vbInformation + vbSystemModal, "Exemptions for EATS Household members with MEMB 01"

    End If
end function

function TLR_determine_SNAP_unit(eats_group_members, memb_found, eats_HH_count)
    eats_group_members = ""
    memb_found = True
    eats_HH_count = 0

    CALL navigate_to_MAXIS_screen("STAT", "EATS")
    EMReadScreen all_eat_together, 1, 4, 72

    IF all_eat_together = "_" Then                          'single member HH's
        eats_group_members = "01" & ","
		eats_HH_count = 1
    ELSEIF all_eat_together = "Y" THEN                      'HH's where all members eat together
        eats_row = 5
        DO
            EMReadScreen eats_pers, 2, eats_row, 3
            eats_pers = replace(eats_pers, " ", "")
            IF trim(eats_pers) = "" THEN
                Exit do
            Else
                eats_group_members = eats_group_members & eats_pers & ","
				eats_HH_count = eats_HH_count  + 1
                eats_row = eats_row + 1
            END IF
        LOOP
    ELSEIF all_eat_together = "N" Then                      'multiple eats HH cases - only eval the 1st eats group that contains MEMB 01.
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
end function

function TLR_active_progs_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, snap_status)
    'Case-based determination
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

    For items = 0 to UBound(eats_group_array, 2)
        '----------------------------------------------------------------------------------------------------14 – ES Compliant While Receiving MFIP
        If mfip_case = True then
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "MFIP Active. "
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "14" & "|"
        End if
    Next

    '----------------------------------------------------------------------------------------------------17 – Receiving RCA
	'Person-based determination -- Looking for RCA information while still on CASE/CURR
	row = 1
    col = 1
    EMSearch "RCA:", row, col
    If row <> 0 Then
		EMReadScreen rca_status, 9, row, col + 5
        rca_status = trim(rca_status)
		rca_status = rca_status
        If rca_status = "ACTIVE" or rca_status = "APP CLOSE" or rca_status = "APP OPEN" Then
			'Navigate to ELIG/RCA to verify if member is eligible for RCA
			EMWriteScreen "ELIG", 20, 22
			CALL write_value_and_transmit("RCA ", 20, 69)

			EMReadScreen no_RCA, 10, 24, 2
			If no_RCA <> "NO VERSION" then
				'RCA version exists so should eb at ELIG/RCA now
				EMWriteScreen "99", 19, 78
				transmit
				'This brings up the FS versions of eligibility results to search for approved versions
				status_row = 7
				Do
					EMReadScreen app_status, 8, status_row, 50
					app_status = trim(app_status)
					If app_status = "" then
						PF3
						exit do 	'if end of the list is reached then exits the do loop
					End if
					If app_status = "UNAPPROV" Then status_row = status_row + 1
				Loop until app_status = "APPROVED" or app_status = ""

				If app_status = "APPROVED" then
					EMReadScreen vers_number, 1, status_row, 23
					Call write_value_and_transmit(vers_number, 18, 54)

					'Read the status for all HH membs
					For items = 0 to UBound(eats_group_array, 2)
						'Read the Elig Status for each HH Member
						status_row = 7
						Do
							EMReadScreen ref_number, 2, status_row, 6
							ref_number = trim(ref_number)
							If ref_number = "" then
								'Check if we are on last page of members - try to PF8 to next page
								PF8
								EMReadScreen members_display_check, 10, 24, 2
								If members_display_check = "** NO MORE" Then
									'Last page reached without finding matching HH memb, reset back to first page for next member
									Do
										PF7
										EmReadScreen first_page_check, 20, 24, 2
									Loop until first_page_check = "** THIS IS THE FIRST"
									exit do		'Exit do to move to next HH memb
								Else
									'If script successfully navigated to next page then status_row needs to be reset
									status_row = 7
								End If
							ElseIf ref_number = eats_group_array(memb_number_const, items) then
								'Found the matching Ref Number, check on Elig Status
								EmReadScreen elig_status, 10, status_row, 53
								elig_status = trim(elig_status)
								If elig_status = "ELIGIBLE" Then
									eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "RCA Active and Eligible. "
									eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "17" & "|"
								End If

								'Regardless of whether HH member is eligible for RCA, need to reset back to start to search next HH memb
								Do
									PF7
									EmReadScreen first_page_check, 20, 24, 2
								Loop until first_page_check = "** THIS IS THE FIRST"
								exit do 	'Exit do to move to next HH memb
							Else
								'If no match found, then move to the next row
								status_row = status_row + 1
							End If
						Loop until ref_number = eats_group_array(memb_number_const, items)
					Next
				End If
			End If
        End If
	End if
end function


function TLR_demographic_exemptions(eats_group_members, eats_group_array, memb_number_const, memb_name_const, verified_exemption_const, verified_wreg_const, potential_exempt_const, meets_childcare_exemption)

    child_under_six = False 	'defaulting to False
	child_under_14 = False		'defaulting to False
	adult_HH_count = 0
    meets_childcare_exemption = True

    'person-based determination (age-based exemptions): STAT/MEMB
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")

    For cow = 0 to UBound(eats_group_members)
    ' For items = 0 to UBound(eats_group_array, 2)
        cl_age = ""
        CALL write_value_and_transmit(eats_group_members(cow), 20, 76)
        ' CALL write_value_and_transmit(eats_group_array(memb_number_const, items), 20, 76)
        EMReadScreen first_name, 12, 6, 63
        first_name = replace(first_name, "_", "")
        Call fix_case_for_name(first_name)
        ' If cow =< UBound(eats_group_array, 2) Then

        EMReadScreen cl_age, 2, 8, 76
        cl_age = trim(cl_age)
        IF cl_age = "" THEN cl_age = 0
        cl_age = cl_age * 1

        If cow =< UBound(eats_group_array, 2) Then
            If eats_group_members(cow) = eats_group_array(memb_number_const, cow) Then
                eats_group_array(memb_name_const, cow) = first_name
                eats_group_array(memb_age_const, cow) = cl_age
            End If
        ElseIf UBound(eats_group_array, 2) = 0 Then
            If eats_group_members(cow) = eats_group_array(memb_number_const, 0) Then
                eats_group_array(memb_name_const, 0) = first_name
                eats_group_array(memb_age_const, 0) = cl_age
            End If
        End If

        'case-based exemption
		If cl_age < 6 then child_under_six = True
        IF cl_age =< 13 THEN
			child_under_14 = True
		Else
			adult_HH_count = adult_HH_count + 1
		End if
		If (child_under_14 = False and child_14_to_17 = True) then meets_childcare_exemption = False

    NEXT

    '----------------------------------------------------------------------------------------------------21 – Child < 18 Living in the SNAP Unit
    For items = 0 to UBound(eats_group_array, 2)
        CALL write_value_and_transmit(eats_group_array(memb_number_const, items), 20, 76)

        native = False
        EMReadScreen tribal_indicator, 2, 18, 42
        EmReadScreen race_detail, 37, 17, 42
        If trim(race_detail) = "Amer Indn Or Alaskan Native" then native = True
        If trim(race_detail) = "Unable To Determine" then
            eats_group_array(potential_exempt_const, items) = eats_group_array(potential_exempt_const, items) & "No race indicated. "
        End If
        If trim(race_detail) = "Multiple Races" then
            PF9
            Call write_value_and_transmit("X", 17, 34)
            EMReadScreen native_indicator, 1, 10, 12
            If native_indicator = "X" then native = True
            transmit 'to exit pop up
            PF10
			Call MAXIS_background_check
        End if
		If native = true then
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "American Indian or Alaskan Native. "
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "30" & "|"
        End If

        If child_under_14 = True then
            If eats_group_array(memb_age_const, items) > 17 then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Child under 14 in SNAP Household. "
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "21" & "|"
            End if
        End if
        '----------------------------------------------------------------------------------------------------08 – Responsible for care of child <6 years old
        If child_under_six = True then
            If adult_HH_count = 1 then
                If eats_group_array(memb_age_const, items) > 17 then
                    eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Care of child under 6. "
                    eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "08" & "|"
                End if
            Else
                eats_group_array(potential_exempt_const, items) = eats_group_array(potential_exempt_const, items) & "Child under 6 in SNAP Household. "
            End if
        End if
        '----------------------------------------------------------------------------------------------------07 – Age 16-17, Living W/Pare/Crgvr
        If eats_group_array(memb_age_const, items) = 16 or eats_group_array(memb_age_const, items) = 17 then
			EMReadScreen age_verif_code, 2, 8, 68
			If age_verif_code <> "NO" then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 16-17. "
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "07" & "|"
			End if
		End if
		'----------------------------------------------------------------------------------------------------06 – Under age 16
        If eats_group_array(memb_age_const, items) < 16 then
		    If age_verif_code <> "NO" then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Under age 16. "
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "06" & "|"
		    End if
		End if
        '----------------------------------------------------------------------------------------------------'16 – 55-59 Years Old
        If eats_group_array(memb_age_const, items) => 50 then
		    If eats_group_array(memb_age_const, items) =< 59 then
		    	If age_verif_code <> "NO" then
                    eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 50-59. "
                    eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "16" & "|"
		    	End if
		    End if
			If (cl_age => 50 and cl_age =< 55) then age_50_thru_55 = True
		End if
        '----------------------------------------------------------------------------------------------------'05 - Age 60 or older
		If eats_group_array(memb_age_const, items) => 65 then
		    If age_verif_code <> "NO" then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 65 or older. "
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "05" & "|"
			End if
		End if
		If eats_group_array(memb_age_const, items) => 60 then
    		If eats_group_array(memb_age_const, items) =< 64 then
                If age_verif_code <> "NO" then
                    eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 60-64. "
                    eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "05" & "|"
                    age_60_thru_64  = True
    			End if
    		End if
    	End if
    Next



end function


function TLR_disability_exemptions(eats_group_members, eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const)
    disabled_eats_member = False
    Call navigate_to_MAXIS_screen("STAT", "DISA")
    single_memb_ref = ""
    If UBound(eats_group_array, 2)  = 0 Then single_memb_ref = eats_group_array(memb_number_const, 0)

    For cow = 0 to UBound(eats_group_members)
        CALL write_value_and_transmit(eats_group_members(cow), 20, 76)
		verified_disa = False
		disa_status = False

        EMReadScreen num_of_DISA, 1, 2, 78
        If num_of_DISA <> "0" THEN
            EMReadScreen disa_end_dt, 10, 6, 69
            disa_end_dt = replace(disa_end_dt, " ", "/")
            EMReadScreen cert_end_dt, 10, 7, 69
            cert_end_dt = replace(cert_end_dt, " ", "/")
            If IsDate(disa_end_dt) = True THEN
                If DateDiff("D", ABAWD_eval_date, disa_end_dt) > 0 THEN
                    disa_status = True
                    If eats_group_members(cow) <> single_memb_ref then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " - DISA end date = " & disa_end_dt & ". "
                End If
            Else
                If disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
                    disa_status = True
                    If eats_group_members(cow) <> single_memb_ref then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " -DISA has no end date. "
                End If
            End If
            If IsDate(cert_end_dt) = True AND disa_status = False THEN
                If DateDiff("D", ABAWD_eval_date, cert_end_dt) > 0 THEN
                    If eats_group_members(cow) <> single_memb_ref then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " - " & cert_end_dt & ". "
                End if
            Else
                If cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
                    EMReadScreen cert_begin_dt, 8, 7, 47
                    If cert_begin_dt <> "__ __ __" THEN
                        disa_status = True
                        If eats_group_members(cow) <> single_memb_ref then possible_exemptions = possible_exemptions & vbcr & "Appears to have disability exemption for the case of HH member " & eats_pers & " -DISA certification has no end date. "
                    End if
                End If
            End If

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
                                End If
                            End If
                        End If
                    End If
                    row = row + 1
                Loop until row = 14
                If verified_disa = True then
                    If single_memb_ref = "" or single_memb_ref = eats_group_array(memb_number_const, cow) Then
                        eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "03" & "|"
                        eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Disabled. "
                    End If
                End If
            End If
        End If
    Next

end function


function TLR_employed_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, verified_abawd_const)

    For oxen = 0 to UBound(eats_group_array, 2)
        prosp_inc = 0
        prosp_hrs = 0
        prospective_hours = 0
        CALL navigate_to_MAXIS_screen("STAT", "JOBS")
        EMWriteScreen eats_group_array(memb_number_const, oxen), 20, 76
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
        				'Going into the PIC for a job with an end ABAWD_eval_date in the future
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
        CALL write_value_and_transmit(eats_group_array(memb_number_const, oxen), 20, 76)
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
        If num_of_RBIC <> "0" then eats_group_array(potential_exempt_const, oxen) = eats_group_array(potential_exempt_const, oxen) & "Has RBIC panel. Review manually for exemptions. "

        If prosp_inc >= 935.25 OR prospective_hours >= 129 THEN
			If jobs_verif_code <> "N" or jobs_verif_code <> "N" then
				If busi_verif_code <> "_" or busi_verif_code <> "N" then
                    eats_group_array(verified_exemption_const, oxen) = eats_group_array(verified_exemption_const, oxen) & "Employed 30 hours/week or earnings at least = to federal minimum wage x 30/hours per week (935.25/month). "
                    eats_group_array(verified_wreg_const, oxen) = eats_group_array(verified_wreg_const, oxen) & "09" & "|"
				End if
			End if
        ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN
			If jobs_verif_code <> "N" or jobs_verif_code <> "N" then
				If busi_verif_code <> "_" or busi_verif_code <> "N" then
                    eats_group_array(verified_exemption_const, oxen) = eats_group_array(verified_exemption_const, oxen) & "Employed at least 80 hours per month. "
					eats_group_array(verified_abawd_const, oxen) = eats_group_array(verified_abawd_const, oxen) & "06" & "|"
				End if
			End if
        End If
    Next

end function


function TLR_unearned_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, potential_exempt_const)

    Call navigate_to_MAXIS_screen("STAT", "UNEA")
    For cow = 0 to UBound(eats_group_array, 2)
        Call write_value_and_transmit(eats_group_array(memb_number_const, cow), 20, 76)
        EMReadScreen num_of_UNEA, 1, 2, 78
        If num_of_UNEA <> "0" THEN
        	Do
        		EMReadScreen unea_type, 2, 5, 37
        		EMReadScreen unea_end_dt, 8, 7, 68
        		unea_end_dt = replace(unea_end_dt, " ", "/")
        		If IsDate(unea_end_dt) = True THEN
        			If DateDiff("D", ABAWD_eval_date, unea_end_dt) > 0  or unea_end_dt = "__/__/__" THEN
        				If unea_type = "11" then
        					EmReadScreen VA_verif_code, 1, 5, 65
        					If VA_verif_code <> "N" then
                                eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "VA Disability. "
                                eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "03" & "|"
        						Exit do
        					Else
                                eats_group_array(potential_exempt_const, cow) = eats_group_array(potential_exempt_const, cow) & "Appears to have VA disability benefits. "
        					End if
        				Elseif unea_type = "14" then
		    				EmReadScreen UC_verif_code, 1, 5, 65
		    				If UC_verif_code <> "N" then
                                eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Unemployment. "
		    					eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "11" & "|"
		    					Exit do
		    				Else
                                eats_group_array(potential_exempt_const, cow) = eats_group_array(potential_exempt_const, cow) & "Appears to have active unemployment benefits. "
		    				End if
                        End if
        			End If
        		End If
        		transmit
        		EMReadScreen enter_a_valid, 13, 24, 2
        	Loop until enter_a_valid = "ENTER A VALID"
        End If
    Next

end function

function TLR_pending_benefit_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, potential_exempt_const)

    Call navigate_to_MAXIS_screen("STAT", "PBEN")
    For cow = 0 to UBound(eats_group_array, 2)
        Call write_value_and_transmit(eats_group_array(memb_number_const, cow), 20, 76)
		EMReadScreen num_of_PBEN, 1, 2, 78
        If num_of_PBEN <> "0" Then
        	pben_row = 8
        	Do
                EMreadscreen pben_type, 2, pben_row, 24
                If pben_type = "__" Then Exit Do
        	    If pben_type = "12" Then		'UI pending'
        			EMReadScreen pben_disp, 1, pben_row, 77
        			If pben_disp = "A" Or pben_disp = "P" Then
                        eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Unemployment. "
                        eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "11" & "|"
						Exit Do
                    ElseIf pben_disp = "E" Then
                        eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Unemployment. "
                        eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "11" & "|"
                        Exit Do
        			Else
                        eats_group_array(potential_exempt_const, cow) = eats_group_array(potential_exempt_const, cow) & "May have pending, appealing, or eligible Unemployment benefits. "
                        pben_row = pben_row + 1
                    End If
        		Else
        			pben_row = pben_row + 1
        		End If
        	Loop Until pben_row = 12
		End If
    Next

end function

function TLR_preg_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const)

    Call navigate_to_MAXIS_screen("STAT", "PREG")
    For cow = 0 to UBound(eats_group_array, 2)
        Call write_value_and_transmit(eats_group_array(memb_number_const, cow), 20, 76)
		EMReadScreen num_of_PREG, 1, 2, 78
        If num_of_PREG <> "0" Then
            EMReadScreen preg_due_dt, 8, 10, 53
            preg_due_dt = replace(preg_due_dt, " ", "/")
        	EMReadScreen preg_end_dt, 8, 12, 53
            If preg_due_dt <> "__/__/__" Then
				EMReadScreen preg_verif, 1, 6, 75
                If DateDiff("d", ABAWD_eval_date, preg_due_dt) >= 0 AND preg_end_dt = "__ __ __" Then
					If preg_verif <> "_" Then
                        eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Pregnant. "
                        eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "23" & "|"
					End If
				End If
			End If
        End If
    Next

end function

function TLR_homeless_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const, potential_exempt_const)

	homeless_exemption = False
    possible_homeless = False
    Call navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen homeless_code, 1, 10, 43
	EMReadScreen living_situation, 2, 11, 43
    EMReadScreen addr_line_01, 16, 6, 43
    IF homeless_code = "Y" then
		If living_situation = "02" or _
			living_situation = "06" or _
			living_situation = "07" or _
			living_situation = "08" then
			homeless_exemption = True
		Else
            possible_homeless = True
		End if
    End if

    If homeless_exemption = True or possible_homeless = True then
        For cow = 0 to UBound(eats_group_array, 2)
            If homeless_exemption = True then
                eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Homeless. "
                eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "03" & "|"
            Elseif possible_homeless = True then
                eats_group_array(potential_exempt_const, cow) = eats_group_array(potential_exempt_const, cow) & "Case's ADDR is coded Y for homeless but living situation doesn't match. "
            End if
        Next
    End if

end function

function TLR_school_exemptions(eats_group_array, memb_number_const, verified_exemption_const, verified_wreg_const)

    Call navigate_to_MAXIS_screen("STAT", "SCHL")
    For cow = 0 to UBound(eats_group_array, 2)
        Call write_value_and_transmit(eats_group_array(memb_number_const, cow), 20, 76)
        EMReadScreen num_of_SCHL, 1, 2, 78
        If num_of_SCHL = "1" Then
        	EMReadScreen school_status, 1, 6, 40
            EMReadScreen school_verif, 2, 6, 63
            EMReadScreen SNAP_code, 2, 16, 63
        	If school_status = "F" or school_status = "H" Then
                If school_verif = "SC" or school_verif = "OT" Then
                    If  SNAP_code = "01" or _
                        SNAP_code = "02" or _
                        SNAP_code = "04" or _
                        SNAP_code = "05" or _
                        SNAP_code = "06" or _
                        SNAP_code = "07" or _
                        SNAP_code = "09" or _
                        SNAP_code = "10" Then
                        eats_group_array(verified_exemption_const, cow) = eats_group_array(verified_exemption_const, cow) & "Student. "
                        eats_group_array(verified_wreg_const, cow) = eats_group_array(verified_wreg_const, cow) & "12" & "|"
                    End If
                End If
            End If
		End If
    Next

end function


Function update_WREG_after_review(member_number, cl_age, best_wreg_code, best_abawd_code, meets_childcare_exemption, homeless_exemption, report_notes)
    '----------------------------------------------------------------------------------------------------WREG and ABAWD Workarounds/TLR Record Updates
    'Commenting out manual code for exempt cases. Business has asked that staff update their own TLR Exemptions, including the TLR record.
    manual_code = "M"  'default manual code for exemption cases/counted month code
    counted_month = False 'initializing

    age_50_thru_59 = False
    age_50_thru_55 = False 'This boolean is used for case note workaround. If anyone 50-55, we note that in case note the date of TLR is not specific to 11/01/25 implementation date.
    age_60_thru_64 = False


    If cl_age => 50 then
        If cl_age =< 59 then age_50_thru_59 = True
        If cl_age =< 55 then age_50_thru_55 = True
    End if
    '----------------------------------------------------------------------------------------------------'05 - Age 60 or older
    If cl_age => 60 then
        If cl_age =< 64 then
            age_60_thru_64  = True
        End if
    End if

    age_50_thru_59_workaround = False ' initializing
    If age_50_thru_59 = True then
        If best_wreg_code = "16" then
            age_50_thru_59_workaround = True
            counted_month = True
        End if
    End if

    age_60_thru_64_workaround = False ' initializing
    If age_60_thru_64 = True then
        If best_wreg_code = "05" then
            age_60_thru_64_workaround = True
            counted_month = True
        End if
    End if

    If best_wreg_code = "30" then
        If best_abawd_code = "10" then counted_month = True
        if best_abawd_code = "09" then counted_month = False
        If best_abawd_code = "06" then
            manual_code = "F"	'Does NOT count on the tracking record, and will remove any counted months.
            counted_month = True	'Identified as counted so that the TLR record is updated, but again, not counted.
        End if
        If meets_childcare_exemption = False then counted_month = True
    End if

    Call navigate_to_MAXIS_screen("STAT", "WREG")
    Call write_value_and_transmit(member_number, 20, 76)

    banked_month_case = False 'initializing banked month case variable to determine if we need to add notes about banked months to the report.
    If banked_months_available = True then
        If counted_month = True Then
            Call ABAWD_Tracking_Record(abawd_counted_months, member_number, MAXIS_footer_month) 'Count all the ABAWD months
            If abawd_counted_months => 3 then
                If best_wreg_code = "30" then best_abawd_code = "13"
                manual_code = "C"	'Manual banked months code
                banked_month_case = True
            End if
        End if

    If counted_month = True then
        PF9
        EMWriteScreen best_wreg_code, 8, 50
        EMWriteScreen best_abawd_code, 13, 50
        If best_wreg_code = "30" then
            EmWriteScreen "N", 8, 78
        Else
            EMWriteScreen "_", 8, 78
        End if

        'Updating the ATR if the codes are already not updated for the CM
        ATR_updates = array("D",manual_code)
        For each update_code in ATR_updates
            Call write_value_and_transmit("X", 13, 57) 'Pulls up the WREG tracker'
            bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))      'col to search starts at 15, increased by 4 for each footer month
            If MAXIS_footer_year = CM_yr then
                bene_yr_row = 10
            Else
                bene_yr_row = 9
            End if
            EMReadScreen ATR_code, 1, bene_yr_row, bene_mo_col
            'This bit will only update to the manual codes if the month isn't already reflecting that.
            If manual_code = "F" then
                If ATR_code = "E" or ATR_code = "F" then
                    exit for 'F and E are exmept
                Else
                    Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
                    PF3 'to go back to WREG/Panel
                End if
            ELSEIF manual_code = "M" then
                If ATR_code = "X" or ATR_code = "M" then
                    exit for 'X and M are counted months
                Else
                    Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
                    PF3 'to go back to WREG/Panel
                End if
            ELSEIF manual_code = "C" then
                If ATR_code = "C" or ATR_code = "B" then
                    exit for ' C and B are banked months
                Else
                    Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
                End if
            End if
        Next
        'PF3 'to go back to WREG/Panel
    End if

    Call ABAWD_Tracking_Record(abawd_counted_months, member_number, MAXIS_footer_month) 'Count all the ABAWD months
    'banked months used messaging
    If banked_months_available = True then
        If banked_month_case = True then
            If banked_months_count = 1 then report_notes = report_notes & "Using 1st banked month. "
            If banked_months_count => 2 then report_notes = report_notes & "Used " & banked_months_count & " banked months. Assess for closure. "
        End if
    End if

    If (counted_month = True and manual_code = "M") then
        'Only 30/06's meet this the above criteria. All other counted months will have the assess for closure note added.
        If abawd_counted_months => 3 then report_notes = report_notes & "Assess TLR for closure for next month. "
    End if

    transmit ' to save
    EMReadscreen orientation_warning, 7, 24, 2 	'reading for orientation date warning message. This message has been casuing me TROUBLE!!
    If orientation_warning = "WARNING" then transmit
    PF3 'to save and exit to stat/wrap

    'case note workaround
    If (age_50_thru_59_workaround = True or age_60_thru_64_workaround = True) then
        Call navigate_to_MAXIS_screen("CASE", "NOTE")
        EMReadScreen first_case_note, 34, 5, 25

        If first_case_note <> "--SNAP Time Limited Recipient: Age" then
            If age_50_thru_59 = True then
                If age_50_thru_55 = False then
                    TLR_text = "55-59"
                    TLR_coding = "16/03"
                Else
                    TLR_text = "50-54"
                    TLR_coding = "16/03"
                End if
            End if

            If age_60_thru_64 = True then
                TLR_text = "60-64"
                TLR_coding = "05/01"
            End if
            Call start_a_blank_CASE_NOTE
            Call write_variable_in_CASE_NOTE("--SNAP Time Limited Recipient: Age " & cl_age & "--")
            Call write_variable_in_CASE_NOTE("TLR member #" & member_number)
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE("* " & TLR_text & " year olds are no longer exempt from SNAP time limits due solely to age.")
            'TODO Add ER or Application date >= 11/01/2025 case noting here when updating that section of the script.
            Call write_variable_in_CASE_NOTE("* FSET/ABAWD codes continue to be " & TLR_coding & " until DHS system updates are in place. ABAWD Tracking record has been updated for this month as a counted month per policy.")
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE(Worker_Signature)
            PF3
        End if
        report_notes = report_notes & cl_age & " year old! "
    End if

    If homeless_exemption = True then
        Call navigate_to_MAXIS_screen("CASE", "NOTE")
        EMReadScreen first_case_note, 40, 5, 25
        If first_case_note <> "--SNAP Time Limited Exempt: Homelessness" then
            start_a_blank_CASE_NOTE
            Call write_variable_in_CASE_NOTE("--SNAP Time Limited Exempt: Homelessness--")
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE("* Case is code as homeless on ADDR, and has applicable living situation which exempts this case from SNAP Work Rules and time limits.")
            Call write_variable_in_CASE_NOTE("* FSET/ABAWD codes are 03/01 for members whom meet this exemption.")
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE(Worker_Signature)
        End if
        PF3
    End if
End Function

Call script_end_procedure("Did that work?")