'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 98                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

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
call changelog_update("08/19/2019", "Updated script so that if started from the ABAWD Tracking Record pop-up on WREG, the script will read where the cursor is placed in the tracking record and if placed on a specific month, the script will autofill that footer month.", "Casey Love, Hennepin County")
call changelog_update("05/07/2018", "Updated universal ABWAWD function.", "Ilse Ferris, Hennepin County")
call changelog_update("04/25/2018", "Updated SCHL exemption coding.", "Ilse Ferris, Hennepin County")
call changelog_update("04/16/2018", "Updated output of potential exemptions for readability.", "Ilse Ferris, Hennepin County")
call changelog_update("04/10/2018", "Enhanced to check cases coded for homelessness for the 'Unfit for Employment' expansion. Also removed code that checked for SSI applying/appealing as this is no longer an exemption reason.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function ABAWD_FSET_exemption_finder_test()
'excluding matching grant and participating in CD treatment due to non-MAXIS indicators.
'excluding armed forces participation dur to non-MAXIS indicators. 
'----------------------------------------------------------------------------------------------------Determining the EATS Household

    Dim eats_group_array()
    ReDim eats_group_array(memb_verified_abawd_const,0)

    'constants for array
    const memb_name_const           = 0
    const memb_number_const         = 1
    const memb_age_const            = 2
    const verified_exemption_const  = 3
    const potential_exempt_const    = 4
    const verified_wreg_const       = 5
    const verified_abawd_const      = 6
    
    entry_record = 0
    case_based_exemptions = ""
    eats_HH_count = 0

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

    eats_group_members = trim(eats_group_members)
    eats_group_members = split(eats_group_members, ",")

    For each memb in eats_group_members    
    	ReDim Preserve eats_group_array(memb_verified_abawd_const, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    	eats_group_array(memb_number_const, entry_record) = memb
    	entry_record = entry_record + 1			'This increments to the next entry in the array'
    	stats_counter = stats_counter + 1
    Next 

    msgbox entry_entry
    
    'Case-based determination
    '----------------------------------------------------------------------------------------------------14 – ES Compliant While Receiving MFIP
	'----------------------------------------------------------------------------------------------------20 – ES Compliant While Receiving DWP
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
	
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

     For items = 0 to UBound(eats_group_array, 2)    
        If mfip_case = True then 
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "MFIP Active" & "|"
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "14" & "|"
        End if 

        If DWP_case = True then 
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "DWP Active" & "|"
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "20" & "|"
        End if 
    
        If rca_case = TRUE = True then 
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "RCA Active" & "|"
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "17" & "|"
        End if 
    Next 

	'----------------------------------------------------------------------------------------------------'Foster care on 18th 
	''<<<<<<<<<<PROG for Foster care
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

    '----------------------------------------------------------------------------------------------------Age-Based Exemptions
    child_under_six = False 	'defaulting to False
	child_under_18 = False		'defaulting to False
	adult_HH_count = 0

    'age_50 = False
    'age_53_54 = False 
    'age_53_54_counted = False 'temporary coding to support. Effective 10/1/24 53-54 YO's starting being TLR's after their next renewal
    'If cl_age = 50 or _
	'	cl_age = 51 or _ 
	'	cl_age = 52 then 
	'	age_50 = True
	'End if
    'If cl_age = 53 or _
    '    cl_age = 54 then
    '    age_53_54 = True
    'End if 

    CALL navigate_to_MAXIS_screen("STAT", "MEMB")

    For items = 0 to UBound(eats_group_array, 2)    
        CALL write_value_and_transmit(eats_group_array(memb_number_const, items), 20, 76)
        EMReadScreen, first_name, 12, 6, 63
        first_name = replace(first_name, "_", "")
        Call fix_case_for_name(first_name)
        eats_group_array(memb_name_const, items) = first_name

        EMReadScreen cl_age, 2, 8, 76
        cl_age = trim(cl_age)
        IF cl_age = "" THEN cl_age = 0
        cl_age = cl_age * 1
        eats_group_array(memb_age_const, items) = cl_age

        'case-based exemption 
		If cl_age < 6 then child_under_six = True
        IF cl_age =< 17 THEN
			child_under_18 = True
		Else
			adult_HH_count = adult_HH_count + 1
		End if
    NEXT

    'person-based determination (age-based exemptions)
    '----------------------------------------------------------------------------------------------------21 – Child < 18 Living in the SNAP Unit
    For items = 0 to UBound(eats_group_array, 2)   
        If child_under_18 = True then 
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Child under 18 in SNAP Household." & "|"
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "21" & "|"
        End if 
        '----------------------------------------------------------------------------------------------------08 – Responsible for care of child <6 years old
        If child_under_6 = True then
            If adult_HH_count = 1 then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Care of child under 6." & "|"
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "08" & "|"
            Else     
                eats_group_array(potential_exempt_const, items) = eats_group_array(potential_exempt_const, items) & "Child under 6 in SNAP Household." & "|"
            End if 
        End if 
        '----------------------------------------------------------------------------------------------------07 – Age 16-17, Living W/Pare/Crgvr
        If cl_age = 16 or cl_age = 17 then
			EMReadScreen age_verif_code, 2, 8, 68
			If age_verif_code <> "NO" then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 16-17". & "|"
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "07" & "|"
			End if
		End if
		'----------------------------------------------------------------------------------------------------06 – Under age 16
        If cl_age < 16 then
		    If age_verif_code <> "NO" then
                eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Under age 16.". & "|"
                eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "06" & "|"
		    End if
		End if
        '----------------------------------------------------------------------------------------------------'16 – 55-59 Years Old
        If cl_age => 55 then
		    If cl_age < 60 then
		    	If age_verif_code <> "NO" then
                    eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 55-59.". & "|"
                    eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "16" & "|"
		    	End if
		    End if
		End if
        '----------------------------------------------------------------------------------------------------'05 - Age 60 or older
		If cl_age => 60 then
		If age_verif_code <> "NO" then
            eats_group_array(verified_exemption_const, items) = eats_group_array(verified_exemption_const, items) & "Age 60 or older.". & "|"
            eats_group_array(verified_wreg_const, items) = eats_group_array(verified_wreg_const, items) & "05" & "|"
			End if
		End if

        '----------------------------------------------------------------------------------------------------possible exemption for foster care members under 24 YO. 
		If cl_age < 24 then 
			If foster_care = True then eats_group_array(potential_exempt_const, items) = eats_group_array(potential_exempt_const, items) & "Under age 24 & may have been in foster case on 18th birthday. Review for exemption. "
		End if 
    Next 