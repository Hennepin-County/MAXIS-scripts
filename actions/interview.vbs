'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
BeginDialog Dialog1, 0, 0, 550, 385, "Dialog"
  GroupBox 180, 5, 300, 260, "Client Conversation"
  GroupBox 10, 5, 170, 260, "Address Information listed in MAXIS"
  ButtonGroup ButtonPressed
    PushButton 50, 245, 125, 15, "CAF has Different Information", caf_info_different_btn
    PushButton 485, 10, 60, 15, "CAF Page 1", caf_page_one_btn
    PushButton 415, 365, 50, 15, "NEXT", next_btn
    PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
EndDialog

class mx_hh_member

	public access_denied
	public selected
	'stuff about the members
	public first_name
	public last_name
	public mid_initial
	public date_of_birth
	public age
	public ref_number
	public ssn
	public ssn_verif
	public birthdate_verif
	public gender
	public id_verif
	public rel_to_applcnt
	public race
	public snap_minor
	public cash_minor
	public written_lang
	public spoken_lang

	public marital_status
	public spouse_ref
	public spouse_name
	public last_grade_completed
	public citizen
	public other_st_FS_end_date
	public in_mn_12_mo
	public residence_verif
	public mn_entry_date
	public former_state

	public left_hh_date

	public imig_exists
	public imig_status
	public us_entry_date
	public imig_status_date
	public imig_status_verif
	public lpr_adj_from
	public nationality
	public alien_id_nbr

	public disa_exists
	public disa_begin_date
	public disa_end_date
	public disa_cert_begin_date
	public disa_cert_end_date
	public cash_disa_status
	public cash_disa_verif
	public fs_disa_status
	public fs_disa_verif
	public hc_disa_status
	public hc_disa_verif
	public disa_waiver
	public disa_1619

	public fs_pwe
	public wreg_exists

	public checkbox_one
	public checkbox_two
	public checkbox_three
	public checkbox_four

	public detail_one
	public detail_two
	public detail_three
	public detail_four

	public button_one
	public button_two
	public button_three
	public button_four

	public property get full_name
		full_name = first_name & " " & last_name
	end property

	Public sub define_the_member()
		Call navigate_to_MAXIS_screen("STAT", "MEMB")
		EMWriteScreen ref_number, 20, 76
		transmit

		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			last_name = "UNABLE TO FIND"
			first_name = "Access Denied"
			mid_initial = ""
			access_denied = TRUE
		Else
			access_denied = FALSE
			EMReadscreen last_name, 25, 6, 30
			EMReadscreen first_name, 12, 6, 63
			EMReadscreen mid_initial, 1, 6, 79
			EMReadScreen age, 3, 8, 76

			EMReadScreen date_of_birth, 10, 8, 42
			EMReadScreen ssn, 11, 7, 42
			EMReadScreen ssn_verif, 1, 7, 68
			EMReadScreen birthdate_verif, 2, 8, 68
			EMReadScreen gender, 1, 9, 42
			EMReadScreen race, 30, 17, 42
			EMReadScreen spoken_lang, 20, 12, 42
			EMReadScreen written_lang, 29, 13, 42


			age = trim(age)
			If age = "" Then age = 0
			age = age * 1
			last_name = trim(replace(last_name, "_", ""))
			first_name = trim(replace(first_name, "_", ""))
			mid_initial = replace(mid_initial, "_", "")
			EMReadScreen id_verif, 2, 9, 68

			EMReadScreen rel_to_applcnt, 2, 10, 42              'reading the relationship from MEMB'
			If rel_to_applcnt = "01" Then rel_to_applcnt = "Self"
			If rel_to_applcnt = "02" Then rel_to_applcnt = "Spouse"
			If rel_to_applcnt = "03" Then rel_to_applcnt = "Child"
			If rel_to_applcnt = "04" Then rel_to_applcnt = "Parent"
			If rel_to_applcnt = "05" Then rel_to_applcnt = "Sibling"
			If rel_to_applcnt = "06" Then rel_to_applcnt = "Step Sibling"
			If rel_to_applcnt = "08" Then rel_to_applcnt = "Step Child"
			If rel_to_applcnt = "09" Then rel_to_applcnt = "Step Parent"
			If rel_to_applcnt = "10" Then rel_to_applcnt = "Aunt"
			If rel_to_applcnt = "11" Then rel_to_applcnt = "Uncle"
			If rel_to_applcnt = "12" Then rel_to_applcnt = "Niece"
			If rel_to_applcnt = "13" Then rel_to_applcnt = "Nephew"
			If rel_to_applcnt = "14" Then rel_to_applcnt = "Cousin"
			If rel_to_applcnt = "15" Then rel_to_applcnt = "Grandparent"
			If rel_to_applcnt = "16" Then rel_to_applcnt = "Grandchild"
			If rel_to_applcnt = "17" Then rel_to_applcnt = "Other Relative"
			If rel_to_applcnt = "18" Then rel_to_applcnt = "Legal Guardian"
			If rel_to_applcnt = "24" Then rel_to_applcnt = "Not Related"
			If rel_to_applcnt = "25" Then rel_to_applcnt = "Live-in Attendant"
			If rel_to_applcnt = "27" Then rel_to_applcnt = "Unknown"

			If id_verif = "BC" Then id_verif = "BC - Birth Certificate"
			If id_verif = "RE" Then id_verif = "RE - Religious Record"
			If id_verif = "DL" Then id_verif = "DL - Drivers License/ST ID"
			If id_verif = "DV" Then id_verif = "DV - Divorce Decree"
			If id_verif = "AL" Then id_verif = "AL - Alien Card"
			If id_verif = "AD" Then id_verif = "AD - Arrival//Depart"
			If id_verif = "DR" Then id_verif = "DR - Doctor Stmt"
			If id_verif = "PV" Then id_verif = "PV - Passport/Visa"
			If id_verif = "OT" Then id_verif = "OT - Other Document"
			If id_verif = "NO" Then id_verif = "NO - No Veer Prvd"

			If age > 18 then
				cash_minor = FALSE
			Else
				cash_minor = TRUE
			End If
			If age > 21 then
				snap_minor = FALSE
			Else
				snap_minor = TRUE
			End If

			date_of_birth = replace(date_of_birth, " ", "/")
			If birthdate_verif = "BC" Then birthdate_verif = "BC - Birth Certificate"
			If birthdate_verif = "RE" Then birthdate_verif = "RE - Religious Record"
			If birthdate_verif = "DL" Then birthdate_verif = "DL - Drivers License/State ID"
			If birthdate_verif = "DV" Then birthdate_verif = "DV - Divorce Decree"
			If birthdate_verif = "AL" Then birthdate_verif = "AL - Alien Card"
			If birthdate_verif = "DR" Then birthdate_verif = "DR - Doctor Statement"
			If birthdate_verif = "OT" Then birthdate_verif = "OT - Other Document"
			If birthdate_verif = "PV" Then birthdate_verif = "PV - Passport/Visa"
			If birthdate_verif = "NO" Then birthdate_verif = "NO - No Verif Provided"

			ssn = replace(ssn, " ", "-")
			if ssn = "___-__-____" Then ssn = ""
			If ssn_verif = "A" THen ssn_verif = "A - SSN Applied For"
			If ssn_verif = "P" THen ssn_verif = "P - SSN Provided, verif Pending"
			If ssn_verif = "N" THen ssn_verif = "N - SSN Not Provided"
			If ssn_verif = "V" THen ssn_verif = "V - SSN Verified via Interface"

			If gender = "M" Then gender = "Male"
			If gender = "F" Then gender = "Female"

			race = trim(race)

			spoken_lang = replace(replace(spoken_lang, "_", ""), "  ", " - ")
			written_lang = trim(replace(replace(replace(written_lang, "_", ""), "  ", " - "), "(HRF)", ""))
		End If

		If access_denied = FALSE Then
			Call navigate_to_MAXIS_screen("STAT", "MEMI")
			EMWriteScreen ref_number, 20, 76
			transmit

			EMReadScreen marital_status, 1, 7, 40
			EMReadScreen spouse_ref, 2, 9, 49
			EMReadScreen spouse_name, 40, 9, 52
			EMReadScreen last_grade_completed, 2, 10, 49
			EMReadScreen citizen, 1, 11, 49
			EMReadScreen other_st_FS_end_date, 8, 13, 49
			EMReadScreen in_mn_12_mo, 1, 14, 49
			EMReadScreen residence_verif, 1, 14, 78
			EMReadScreen mn_entry_date, 8, 15, 49
			EMReadScreen former_state, 2, 15, 78

			If marital_status = "N" Then marital_status = "N - Never Married"
			If marital_status = "M" Then marital_status = "M - Married Living with Spouse"
			If marital_status = "S" Then marital_status = "S - Married Living Apart"
			If marital_status = "L" Then marital_status = "L - Legally Seperated"
			If marital_status = "D" Then marital_status = "D - Divorced"
			If marital_status = "W" Then marital_status = "W - Widowed"
			If spouse_ref = "__" Then spouse_ref = ""
			spouse_name = trim(spouse_name)

			If last_grade_completed = "00" Then last_grade_completed = "Not Attended or Pre-Grade 1 - 00"
			If last_grade_completed = "12" Then last_grade_completed = "High School Diploma or GED - 12"
			If last_grade_completed = "13" Then last_grade_completed = "Some Post Sec Education - 13"
			If last_grade_completed = "14" Then last_grade_completed = "High School Plus Certiificate - 14"
			If last_grade_completed = "15" Then last_grade_completed = "Four Year Degree - 15"
			If last_grade_completed = "16" Then last_grade_completed = "Grad Degree - 16"
			If len(last_grade_completed) = 2 Then last_grade_completed = "Grade " & last_grade_completed
			If citizen = "Y" Then citizen = "Yes"
			If citizen = "N" Then citizen = "No"

			other_st_FS_end_date = replace(other_st_FS_end_date, " ", "/")
			If other_st_FS_end_date = "__/__/__" Then other_st_FS_end_date = ""
			If in_mn_12_mo = "Y" Then in_mn_12_mo = "Yes"
			If in_mn_12_mo = "N" Then in_mn_12_mo = "No"
			If residence_verif = "1" Then residence_verif = "1 - Rent Receipt"
			If residence_verif = "2" Then residence_verif = "2 - Landlord's Statement"
			If residence_verif = "3" Then residence_verif = "3 - Utility Bill"
			If residence_verif = "4" Then residence_verif = "4 - Other"
			If residence_verif = "N" Then residence_verif = "N - Verif Not Provided"
			mn_entry_date = replace(mn_entry_date, " ", "/")
			If mn_entry_date = "__/__/__" Then mn_entry_date = ""
			If former_state = "__" Then former_state = ""


			Call navigate_to_MAXIS_screen("STAT", "IMIG")
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen imig_version, 1, 2, 73
			If imig_version = "0" Then imig_exists = FALSE
			If imig_version = "1" Then imig_exists = TRUE

			If imig_exists = TRUE Then
				EMReadScreen imig_status, 40, 6, 45
				EMReadScreen us_entry_date, 10, 7, 45
				EMReadScreen imig_status_date, 10, 7, 71
				EMReadScreen imig_status_verif, 2, 8, 45
				EMReadScreen lpr_adj_from, 40, 9, 45
				EMReadScreen nationality, 2, 10, 45
				EMReadScreen alien_id_nbr, 10, 10, 71

				imig_status = trim(imig_status)
				us_entry_date = replace(us_entry_date, " ", "/")
				imig_status_date = replace(imig_status_date, " ", "/")

				If imig_status_verif = "S1" Then imig_status_verif = "S1 - SAVE Primary"
				If imig_status_verif = "S2" Then imig_status_verif = "S2 - SAVE Secondary"
				If imig_status_verif = "AL" Then imig_status_verif = "AL - Alien Card"
				If imig_status_verif = "PV" Then imig_status_verif = "PV - Passport/Visa"
				If imig_status_verif = "RE" Then imig_status_verif = "RE - Re-Entry Permit"
				If imig_status_verif = "IM" Then imig_status_verif = "IN - INS Correspondence"
				If imig_status_verif = "OT" Then imig_status_verif = "OT - Other Document"
				If imig_status_verif = "NO" Then imig_status_verif = "NO - No Verif Provided"

				lpr_adj_from = trim(lpr_adj_from)

				If nationality = "AA" Then nationality = "AA - Amerasian"
				If nationality = "EH" Then nationality = "EH - Ethnic Chinese"
				If nationality = "EL" Then nationality = "EL - Ethnic Lao"
				If nationality = "HG" Then nationality = "HG - Hmong"
				If nationality = "KD" Then nationality = "KD - Kurd"
				If nationality = "SJ" Then nationality = "SJ - Soviet Jew"
				If nationality = "TT" Then nationality = "TT - Tinh"
				If nationality = "AF" Then nationality = "AF - Afghanistan"
				If nationality = "BK" Then nationality = "BK - Bosnia"
				If nationality = "CB" Then nationality = "CB - Cambodia"
				If nationality = "CH" Then nationality = "CH - China, Mainland"
				If nationality = "CU" Then nationality = "CU - Cuba"
				If nationality = "ES" Then nationality = "ES - El Salvador"
				If nationality = "ER" Then nationality = "ER - Eritrea"
				If nationality = "ET" Then nationality = "ET - Ethiopia"
				If nationality = "GT" Then nationality = "GT - Guatemala"
				If nationality = "HA" Then nationality = "HA - Haiti"
				If nationality = "HO" Then nationality = "HO - Honduras"
				If nationality = "IR" Then nationality = "IR - Iran"
				If nationality = "IZ" Then nationality = "IZ - Iraq"
				If nationality = "LI" Then nationality = "LI - Liberia"
				If nationality = "MC" Then nationality = "MC - Micronesia"
				If nationality = "MI" Then nationality = "MI - Marshall Islands"
				If nationality = "MX" Then nationality = "MX - Mexico"
				If nationality = "WA" Then nationality = "WA - Namibia (SW Africa)"
				If nationality = "PK" Then nationality = "PK - Pakistan"
				If nationality = "RP" Then nationality = "RP - Philippines"
				If nationality = "PL" Then nationality = "PL - Poland"
				If nationality = "RO" Then nationality = "RO - Romania"
				If nationality = "RS" Then nationality = "RS - Russia"
				If nationality = "SO" Then nationality = "SO - Somalia"
				If nationality = "SF" Then nationality = "SF - South Africa"
				If nationality = "TH" Then nationality = "TH - Thailand"
				If nationality = "VM" Then nationality = "VM - Vietnam"
				If nationality = "OT" Then nationality = "OT - All Others"
			End If


			Call navigate_to_MAXIS_screen("STAT", "DISA")
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen disa_version, 1, 2, 73
			If disa_version = "0" Then disa_exists = FALSE
			If disa_version = "1" Then disa_exists = TRUE

			If disa_exists = TRUE Then
				EMReadScreen disa_begin_date, 10, 6, 47
				EMReadScreen disa_end_date, 10, 6, 69
				EMReadScreen disa_cert_begin_date, 10, 7, 47
				EMReadScreen disa_cert_end_date, 10, 7, 69
				EMReadScreen cash_disa_status, 2, 11, 59
				EMReadScreen cash_disa_verif, 1, 11, 69
				EMReadScreen fs_disa_status, 2, 12, 59
				EMReadScreen fs_disa_verif, 1, 12, 69
				EMReadScreen hc_disa_status, 2, 13, 59
				EMReadScreen hc_disa_verif, 1, 13, 69
				EMReadScreen disa_waiver, 1, 14, 59
				EMReadScreen disa_1619, 1, 16, 59

				disa_begin_date = replace(disa_begin_date, " ", "/")
				If disa_begin_date = "__/__/____" Then disa_begin_date = ""
				disa_end_date = replace(disa_end_date, " ", "/")
				If disa_end_date = "__/__/____" Then disa_end_date = ""
				disa_cert_begin_date = replace(disa_cert_begin_date, " ", "/")
				If disa_cert_begin_date = "__/__/____" Then disa_cert_begin_date = ""
				disa_cert_end_date = replace(disa_cert_end_date, " ", "/")
				If disa_cert_end_date = "__/__/____" Then disa_cert_end_date = ""

				If cash_disa_status = "01" Then cash_disa_status = "01 - RSDI Only Disability"
				If cash_disa_status = "02" Then cash_disa_status = "02 - RSDI Only Blindness"
				If cash_disa_status = "03" Then cash_disa_status = "03 - SSI, SSI/RSDI Disability"
				If cash_disa_status = "04" Then cash_disa_status = "04 - SSI, SSI/RSDI Blindness"
				If cash_disa_status = "06" Then cash_disa_status = "06 - SMRT/SSA Pend"
				If cash_disa_status = "08" Then cash_disa_status = "08 - SMRT Certified Blindness"
				If cash_disa_status = "09" Then cash_disa_status = "09 - Ill/Incapacity"
				If cash_disa_status = "10" Then cash_disa_status = "10 - SMRT Certified Disability"
				If cash_disa_status = "__" Then cash_disa_status = ""
				If cash_disa_verif = "1" Then cash_disa_verif = "1 - DHS 161/Dr Stmt"
				If cash_disa_verif = "2" Then cash_disa_verif = "2 - SMRT Certified"
				If cash_disa_verif = "3" Then cash_disa_verif = "3 - Certified for RSDI or SSI"
				If cash_disa_verif = "6" Then cash_disa_verif = "6 - Other Document"
				If cash_disa_verif = "7" Then cash_disa_verif = "7 - Professional Stmt of Need"
				If cash_disa_verif = "N" Then cash_disa_verif = "N - No Verif Provided"

				If fs_disa_status = "01" Then fs_disa_status = "01 - RSDI Only Disability"
				If fs_disa_status = "02" Then fs_disa_status = "02 - RSDI Only Blindness"
				If fs_disa_status = "03" Then fs_disa_status = "03 - SSI, SSI/RSDI Disability"
				If fs_disa_status = "04" Then fs_disa_status = "04 - SSI, SSI/RSDI Blindness"
				If fs_disa_status = "08" Then fs_disa_status = "08 - SMRT Certified Blindness"
				If fs_disa_status = "09" Then fs_disa_status = "09 - Ill/Incapacity"
				If fs_disa_status = "10" Then fs_disa_status = "10 - SMRT Certified Disability"
				If fs_disa_status = "11" Then fs_disa_status = "11 - VA Determined Pd - 100% Disa"
				If fs_disa_status = "12" Then fs_disa_status = "12 - VA (Other Accept Disa)"
				If fs_disa_status = "13" Then fs_disa_status = "13 - Certified RR Retirement Disa"
				If fs_disa_status = "14" Then fs_disa_status = "14 - Other Govt Permanent Disa"
				If fs_disa_status = "15" Then fs_disa_status = "15 - Disability from MINE List"
				If fs_disa_status = "16" Then fs_disa_status = "16 - Unable to Prepare Purch Own Meal"
				If fs_disa_status = "__" Then fs_disa_status = ""
				If fs_disa_verif = "1" Then fs_disa_verif = "1 - DHS 161/Dr Stmt"
				If fs_disa_verif = "2" Then fs_disa_verif = "2 - SMRT Certified"
				If fs_disa_verif = "3" Then fs_disa_verif = "3 - Certified for RSDI or SSI"
				If fs_disa_verif = "4" Then fs_disa_verif = "4 - Receipt of HC for Disa/Blind"
				If fs_disa_verif = "5" Then fs_disa_verif = "5 - Work Judgement"
				If fs_disa_verif = "6" Then fs_disa_verif = "6 - Other Document"
				If fs_disa_verif = "7" Then fs_disa_verif = "7 - Out of State Verif Pending"
				If fs_disa_verif = "N" Then fs_disa_verif = "N - No Verif Provided"

				If hc_disa_status = "01" Then hc_disa_status = "01 - RSDI Only Disability"
				If hc_disa_status = "02" Then hc_disa_status = "02 - RSDI Only Blindness"
				If hc_disa_status = "03" Then hc_disa_status = "03 - SSI, SSI/RSDI Disability"
				If hc_disa_status = "04" Then hc_disa_status = "04 - SSI, SSI/RSDI Blindness"
				If hc_disa_status = "06" Then hc_disa_status = "06 - SMRT Pend or SSA Pend"
				If hc_disa_status = "08" Then hc_disa_status = "08 - Certified Blind"
				If hc_disa_status = "10" Then hc_disa_status = "10 - Certified Disabled"
				If hc_disa_status = "11" Then hc_disa_status = "11 - Special Category - Disabled Child"
				If hc_disa_status = "20" Then hc_disa_status = "20 - TEFRA - Disabled"
				If hc_disa_status = "21" Then hc_disa_status = "21 - TEFRA - Blind"
				If hc_disa_status = "22" Then hc_disa_status = "22 - MA-EPD"
				If hc_disa_status = "23" Then hc_disa_status = "23 - MA/Waiver"
				If hc_disa_status = "24" Then hc_disa_status = "24 - SSA/SMRT Appeal Pending"
				If hc_disa_status = "26" Then hc_disa_status = "26 - SSA/SMRT Disa Deny"
				If hc_disa_status = "__" Then hc_disa_status = ""
				If hc_disa_verif = "1" Then hc_disa_verif = "1 - DHS 161/Dr Stmt"
				If hc_disa_verif = "2" Then hc_disa_verif = "2 - SMRT Certified"
				If hc_disa_verif = "3" Then hc_disa_verif = "3 - Certified for RSDI or SSI"
				If hc_disa_verif = "6" Then hc_disa_verif = "6 - Other Document"
				If hc_disa_verif = "7" Then hc_disa_verif = "7 - Case Manager Determination"
				If hc_disa_verif = "8" Then hc_disa_verif = "8 - LTC Consult Services"
				If hc_disa_verif = "N" Then hc_disa_verif = "N - No Verif Provided"

				If disa_waiver = "F" Then disa_waiver = "F - LTC CADI Conversion"
				If disa_waiver = "G" Then disa_waiver = "G - LTC CADI DIversion"
				If disa_waiver = "H" Then disa_waiver = "H - LTC CAC Conversion"
				If disa_waiver = "I" Then disa_waiver = "I - LTC CAC Diversion"
				If disa_waiver = "J" Then disa_waiver = "J - LTC EW Conversion"
				If disa_waiver = "K" Then disa_waiver = "K - LTC EW Diversion"
				If disa_waiver = "L" Then disa_waiver = "L - LTC TBI NF Conversion"
				If disa_waiver = "M" Then disa_waiver = "M - LTC TBI NF Diversion"
				If disa_waiver = "P" Then disa_waiver = "P - LTC TBI NB Conversion"
				If disa_waiver = "Q" Then disa_waiver = "Q - LTC TBI NB Diversion"
				If disa_waiver = "R" Then disa_waiver = "R - DD Conversion"
				If disa_waiver = "S" Then disa_waiver = "S - DD Conversion"
				If disa_waiver = "Y" Then disa_waiver = "Y - CSG Conversion"
				If disa_waiver = "_" Then disa_waiver = ""

				If disa_1619 = "A" Then disa_1619 = "A - 1619A Status"
				If disa_1619 = "B" Then disa_1619 = "B - 1619B Status"
				If disa_1619 = "N" Then disa_1619 = "N - No 1619 Status"
				If disa_1619 = "T" Then disa_1619 = "T - 1619 Status Terminated"
				If disa_1619 = "_" Then disa_1619 = ""
			End If

			Call navigate_to_MAXIS_screen("STAT", "WREG")
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen wreg_version, 1, 2, 73
			If wreg_version = "0" Then wreg_exists = FALSE
			If wreg_version = "1" Then wreg_exists = TRUE

			If wreg_exists = TRUE Then
				EMReadScreen wreg_pwe, 1, 6, 68

				If wreg_pwe = "Y" Then fs_pwe = "Yes"
				If wreg_pwe = "N" OR wreg_pwe = "_" Then fs_pwe = "No"
			End If
		End If
	end sub

	Public sub choose_the_members()

	end sub

	' private sub Class_Initialize()
	' end sub
end class


class client_income

	'about the income
	public member

end class


Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(0)

const rela_clt_one_ref		= 0
const rela_clt_two_ref 		= 1
const rela_clt_one_name 	= 2
const rela_clt_two_name 	= 3
const rela_type 			= 4
const verif_req_checkbox 	= 5
const rela_verif			= 6
const rela_pers_one			= 7
const rela_pers_two			= 8
const rela_notes 			= 9

Dim ALL_HH_RELATIONSHIPS_ARRAY()
ReDim ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, 0)

rela_type_dropdown = "Select One..."+chr(9)+"Parent"+chr(9)+"Child"+chr(9)+"Sibling"+chr(9)+"Spouse"+chr(9)+"Grandparent"+chr(9)+"Neice"+chr(9)+"Nephew"+chr(9)+"Aunt"+chr(9)+"Uncle"+chr(9)+"Grandchild"+chr(9)+"Step Parent"+chr(9)+"Step Child"+chr(9)+"Relative Caregiver"+chr(9)+"Foster Child"+chr(9)+"Foster Parent"+chr(9)+"Not Related"+chr(9)+"Legal Guardian"+chr(9)+"Other Relative"+chr(9)+"Cousin"+chr(9)+"Live-in Attendant"+chr(9)+"Unknown"
rela_verif_dropdown = "Type or Select"+chr(9)+"BC - Birth Certificate"+chr(9)+"AR - Adoption Records"+chr(9)+"LG = Legal Guardian"+chr(9)+"RE - Religious Records"+chr(9)+"HR - Hospital Records"+chr(9)+"RP - Recognition of Parentage"+chr(9)+"OT - Other Verifciation"+chr(9)+"NO - No Verif Provided"+chr(9)
memb_droplist = ""
the_pwe_for_this_case = ""

function define_main_dialog()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 550, 385, "Dialog"

	  ButtonGroup ButtonPressed
		If page_display = show_pg_one Then
			Text 495, 12, 60, 13, "CAF Page 1"

			GroupBox 180, 5, 300, 260, "Client Conversation"
			Text 185, 20, 245, 10, "^^2 - Read the Residence Address to the client."
			Text 195, 30, 125, 10, "Ask: Is thatthe address you live at?"
			ComboBox 195, 40, 280, 45, "Select or Type client response"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Household is Homeless", residence_address_response
			PushButton 345, 55, 130, 13, "Record specific changes here", client_verbal_changes_resi_address_btn
			Text 185, 70, 245, 10, "^^3 - Ask: Are you experiencing housing instability (or homelessness)?"
			ComboBox 195, 80, 280, 45, "Select or Type client response"+chr(9)+"Yes"+chr(9)+"No", household_homeless_response
			Text 185, 100, 50, 10, "^^4 - Explain:"
			Text 200, 110, 275, 25, "We use mail as our primary means of communication to let you know if any action is required for any benefits you receive to continue. It is important the address we send mail to has your name listed and that you check it regularly. "
			Text 200, 140, 240, 15, "If a mailing address has been listed on the CAF or in MAXIS - read it to the client.                                         Which Address?"
			DropListBox 365, 150, 110, 45, "Residence Address"+chr(9)+"Mailing Address", which_address_are_we_discussing
			Text 195, 165, 150, 10, "Ask: Can you receive mail at this address?"
			ComboBox 195, 175, 280, 45, "Select or Type client response"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"No - Use General Delivery", mail_received_at_this_address_response
			PushButton 345, 190, 130, 13, "Record specific changes here", client_verbal_changes_mail_address_btn
			Text 185, 205, 140, 10, "^^5 - If General Delivery requested - Explain:"
			Text 195, 215, 260, 25, "GD can be used to have your mail held at the post office. You will need a photo ID to collect mail from GD. As our mail often requires a response within 10 days, you should be checking GD at least every 2-3 days."
			DropListBox 275, 245, 200, 45, "General Delivery Not Requested"+chr(9)+"Explained and Client Confirmed Understanding", confirm_gen_del_explanation
			GroupBox 10, 5, 170, 260, "Address Information listed in MAXIS"
			Text 15, 20, 160, 25, "^^1 - Compare this address to the one entered on the CAF. If there is a difference, press the button below to update."
			Text 20, 60, 70, 10, "Residence Address"
			Text 25, 75, 150, 10, resi_line_one
			If resi_line_two = "" Then
				Text 25, 85, 150, 10, resi_city & ", " & resi_state & " " & resi_zip
				Text 25, 95, 150, 10, "County: " & resi_county

			Else
				Text 25, 85, 150, 10, resi_line_two
				Text 25, 95, 150, 10, resi_city & ", " & resi_state & " " & resi_zip
				Text 25, 105, 150, 10, "County: " & resi_county
			End If
			Text 25, 120, 65, 10, "Homeless: " & homeless
			Text 25, 130, 150, 20, "Living Situation: " & living_sit
			Text 25, 155, 60, 10, "IND RES - " & ind_reservation
			Text 25, 165, 150, 10, "RES NAME - " & res_name
			Text 25, 180, 150, 10, "Verification: " & verif
			If mail_line_one = "" Then
				Text 20, 195, 150, 10, "No MAILING ADDRESS Listed"
			ElseIf mail_line_two = "" Then
				Text 20, 195, 70, 10, "Mailing Address"
				Text 25, 210, 150, 10, mail_line_one
				Text 25, 220, 150, 10, mail_city & ", " & mail_state & " " & mail_zip
			Else
				Text 20, 195, 70, 10, "Mailing Address"
				Text 25, 210, 150, 10, mail_line_one
				Text 25, 220, 150, 10, mail_line_two
				Text 25, 230, 150, 10, mail_city & ", " & mail_state & " " & mail_zip
			End If
			PushButton 50, 245, 125, 15, "CAF has Different Information", caf_info_different_btn
			GroupBox 10, 265, 470, 80, "Phone Contact"
			Text 20, 280, 115, 10, "Current Phone Numbers in MAXIS"
			Text 25, 295, 60, 10, phone_numb_one
			Text 90, 295, 35, 10, phone_type_one
			Text 25, 310, 60, 10, phone_numb_two
			Text 90, 310, 35, 10, phone_type_two
			Text 25, 325, 60, 10, phone_numb_three
			Text 90, 325, 40, 10, phone_type_three
			Text 150, 280, 185, 10, "^^6 - Ask: What is the best phone number to reach out at?"
			EditBox 340, 275, 65, 15, reported_phone_number
			Text 185, 295, 120, 10, "What type of phone number is this?"
			DropListBox 305, 290, 100, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", reported_phone_type
			Text 150, 315, 320, 10, "^^7 - For each number on the lest, Read it to the client and Ask: is this still agood number?"
			PushButton 345, 325, 130, 13, "Record Any Changes to Numbers Here", client_verbal_changes_phone_numbers_btn

			GroupBox 10, 345, 350, 30, "Living Situation"
			Text 20, 360, 125, 10, "^^8 - Ask: What is your living situation?"
			DropListBox 150, 355, 200, 15, "Select One..."+chr(9)+"Own Housing(lease, mortgage, or roomate)"+chr(9)+"Family/Friends due to economic hardship"+chr(9)+"Servc prvdr- foster/group home"+chr(9)+"Hospital/Treatment/Detox/Nursing Home"+chr(9)+"Jail/Prison//Juvenile Det."+chr(9)+"Hotel/Motel"+chr(9)+"Emergency Shelter"+chr(9)+"Place not meant for Housing"+chr(9)+"Declined"+chr(9)+"Unknown", clt_response_living_sit
		End If
		If page_display = show_pg_memb_list Then
			Text 495, 27, 60, 13, "CAF MEMBs"

			' Text 15, 15, 450, 10, "THE QUESTION GOES HERE"
			' Text 135, 35, 65, 10, "Answer on the CAF"
			' Text 265, 35, 70, 10, "Confirm CAF Answer"
			' DropListBox 200, 30, 40, 45, "No"+chr(9)+"Yes", caf_answer
			' ComboBox 340, 30, 125, 45, "", confirm_caf_answer
			grp_box_len = 100 + UBound(HH_MEMB_ARRAY, 1) * 15

			GroupBox 10, 30, 470, grp_box_len, "Household Members Listed in MAXIS"
			GroupBox 410, 70, 60, grp_box_len - 45, "REMO Date"
			Text 15, 45, 240, 10, "^^1 - Ask: Please list everyone that lives at the address/lives with you."
			Text 30, 55, 250, 10, "Check the boxes: Is this member listed on the CAF? and Reported Verbally?"
			Text 20, 75, 35, 10, "Applicant:"
			Text 25, 90, 255, 10, "- MEMB " & HH_MEMB_ARRAY(0).ref_number & "    " & HH_MEMB_ARRAY(0).full_name
			CheckBox 285, 90, 50, 10, "On the CAF", HH_MEMB_ARRAY(0).checkbox_one
			CheckBox 340, 90, 70, 10, "Verbally Reported", HH_MEMB_ARRAY(0).checkbox_two
			EditBox 415, 85, 50, 15, HH_MEMB_ARRAY(0).left_hh_date

			Text 20, 110, 95, 10, "Other Household Members:"
			y_pos = 125
			for i = 1 to UBound(HH_MEMB_ARRAY, 1)
				Text 25, y_pos, 255, 10, "- MEMB " & HH_MEMB_ARRAY(i).ref_number & "    " & HH_MEMB_ARRAY(i).full_name & " - " & HH_MEMB_ARRAY(i).rel_to_applcnt & " of Memb 01"
				CheckBox 285, y_pos, 50, 10, "On the CAF", HH_MEMB_ARRAY(i).checkbox_one
				CheckBox 340, y_pos, 70, 10, "Verbally Reported", HH_MEMB_ARRAY(i).checkbox_two
				EditBox 415, y_pos - 5, 50, 15, HH_MEMB_ARRAY(i).left_hh_date
				y_pos = y_pos + 15
			Next
			PushButton 360, 40, 115, 15, "Add Member not Listed Here", add_memb_to_list_btn
			y_pos = y_pos + 10
			Text 15, y_pos, 345, 10, "^^2 - If any members were not reported, Ask: About each member and if they have left the household."
			y_pos = y_pos + 20
			Text 15, y_pos, 465, 10, "Current AREP in MAXIS: FIRST NAME LAST NAME ALL THE THINGS "
			y_pos = y_pos + 15
			Text 15, y_pos, 185, 10, "^^3 - Check the CAF to see if an Authorized Rep is listed. "
			ComboBox 210, y_pos - 5, 270, 45, " "+chr(9)+"Yes - I would like an AREP"+chr(9)+"No - Do not authorize someone on my case", arep_response
			y_pos = y_pos + 10
			Text 30, y_pos, 245, 10, "Ask: Do you want another person authorized to talk to us about your case?"
			y_pos = y_pos + 20
			Text 15, y_pos, 465, 10, "Current SWKR in MAXIS: FIRST NAME LAST NAME ALL THE THINGS "
			y_pos = y_pos + 15
			Text 15, y_pos, 175, 10, "^^4 - Check the CAF to see if a Social Worker is listed. "
			ComboBox 210, y_pos - 5, 270, 40, " "+chr(9)+"Yes - I would like an AREP"+chr(9)+"No - Do not authorize someone on my case", Combo2
			y_pos = y_pos + 10
			Text 30, y_pos, 275, 10, "Ask: Do you have a social worker you want authorized to talk to us about your case?"
			' PushButton 10, 10, 80, 15, "List of HH Members", hh_list_btn
			Text 20, 13, 80, 15, "List of HH Members"
			PushButton 95, 10, 125, 15, "Review HH Member Information", HH_memb_detail_review
		End If
		If page_display = show_pg_memb_info Then
			Text 495, 27, 60, 13, "CAF MEMBs"

			Text 15, 35, 460, 10, "^^1 - Review the personal information/detail for each household member on this case."
			grp_len = 195
			For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
				If ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) = HH_MEMB_ARRAY(memb_selected).ref_number Then grp_len = grp_len + 20
			Next
			If grp_len = 195 Then grp_len = 215
			GroupBox 50, 65, 425, grp_len, "Information for " & HH_MEMB_ARRAY(memb_selected).full_name
			Text 60, 80, 165, 10, "Name: " & HH_MEMB_ARRAY(memb_selected).full_name
			Text 225, 80, 85, 10, "Review of info response:"
			ComboBox 315, 75, 155, 45, "", review_memb_info_detail
			Text 75, 90, 165, 10, "Age: " & HH_MEMB_ARRAY(memb_selected).age & "      DOB: "  & HH_MEMB_ARRAY(memb_selected).date_of_birth
			Text 75, 100, 235, 10, "Written Language: " & HH_MEMB_ARRAY(memb_selected).written_lang & "       Spoken Language: " & HH_MEMB_ARRAY(memb_selected).spoken_lang
			Text 75, 115, 185, 10, "Proof of Identity:" & HH_MEMB_ARRAY(memb_selected).id_verif
			Text 275, 115, 100, 10, "Identity Proof:"
			ComboBox 330, 115, 135, 45, "Select or Type"+chr(9)+"Not Needed"+chr(9)+"Requested"+chr(9)+"ECF"+chr(9)+"SMI/SOL-Q", identity_proof_found
			CheckBox 80, 125, 140, 10, "Check Here if identity proof is required", id_required_checkbox
			Text 75, 140, 65, 10, "SSN: " & HH_MEMB_ARRAY(memb_selected).ssn
			If left(HH_MEMB_ARRAY(memb_selected).ssn_verif, 1) <> "V" Then Text 155, 140, 195, 10, "SSN has not been validated, review this number with client:"
			If left(HH_MEMB_ARRAY(memb_selected).ssn_verif, 1) <> "V" Then ComboBox 350, 135, 115, 45, "SSN wrong - now updated"+chr(9)+"This number is correct", non_validated_ssn_detail
			If HH_MEMB_ARRAY(memb_selected).spouse_ref <> "" Then Text 60, 155, 400, 10, "Marital Status: " & HH_MEMB_ARRAY(memb_selected).marital_status & "          Spouse's Name: " & HH_MEMB_ARRAY(memb_selected).spouse_ref & " - " & HH_MEMB_ARRAY(memb_selected).spouse_name
			If HH_MEMB_ARRAY(memb_selected).spouse_ref = "" Then Text 60, 155, 350, 10, "Marital Status: " & HH_MEMB_ARRAY(memb_selected).marital_status
			Text 60, 170, 140, 10, "In Minnesota for at least 12 Months: " & HH_MEMB_ARRAY(memb_selected).in_mn_12_mo
			If HH_MEMB_ARRAY(memb_selected).in_mn_12_mo = "No" Then Text 215, 170, 160, 10, "Arrived from: " & HH_MEMB_ARRAY(memb_selected).former_state & "     Arrival Date: " & HH_MEMB_ARRAY(memb_selected).mn_entry_date

			Text 60, 185, 65, 10, "US Citizen: " & HH_MEMB_ARRAY(memb_selected).citizen
			If HH_MEMB_ARRAY(memb_selected).imig_exists = TRUE THen
				Text 65, 200, 410, 10, "Immigration Status: " & HH_MEMB_ARRAY(memb_selected).imig_status & "     Immigration Documentaion: " & HH_MEMB_ARRAY(memb_selected).imig_status_verif
				' Text 200, 200, 250, 10, "Immigration Documentaion: " & HH_MEMB_ARRAY(memb_selected).imig_status_verif
			End If
			Text 60, 215, 175, 10, "Relationships in the Household: This Person is ..."
			Text 295, 215, 160, 10, "Relationship Verification:"
			y_pos = 230
			For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
				If ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) = HH_MEMB_ARRAY(memb_selected).ref_number Then
					DropListBox 75, y_pos, 80, 45,rela_type_dropdown, ALL_HH_RELATIONSHIPS_ARRAY(rela_type, the_rela)
					Text 165, y_pos + 5, 10, 10, "to"
					DropListBox 185, y_pos, 100, 45, "Select One..."&memb_droplist, ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela)
					CheckBox 295, y_pos+ 5, 40, 10, "Required", ALL_HH_RELATIONSHIPS_ARRAY(verif_req_checkbox, the_rela)
					ComboBox 345, y_pos, 110, 45, rela_verif_dropdown & ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, the_rela), ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, the_rela)
					y_pos = y_pos + 20
				End If
			Next
			If y_pos = 230 Then
				Text 75, y_pos, 300, 10, "No relationships between household members known in MAXIS at this time."
				y_pos = y_pos + 20
			End If
			' DropListBox 75, 230, 80, 45, ""+chr(9)+"Parent"+chr(9)+"Child"+chr(9)+"Sibling"+chr(9)+"Spouse"+chr(9)+"Grandparent"+chr(9)+"Neice"+chr(9)+"Nephew"+chr(9)+"Aunt"+chr(9)+"Uncle"+chr(9)+"Grandchild", relationship_type
			' Text 165, 235, 10, 10, "to"
			' DropListBox 185, 230, 100, 45, "", relationship_member
			' CheckBox 295, 230, 40, 10, "Required", verif_req_checkbox
			' ComboBox 345, 230, 110, 45, "", relationship_verif_type
			' DropListBox 75, 250, 80, 45, "", List3
			' Text 165, 250, 10, 10, "to"
			' DropListBox 185, 250, 100, 45, "", List4
			' CheckBox 295, 250, 40, 10, "Required", Check3
			' ComboBox 345, 250, 110, 45, "", Combo5

			PushButton 75, y_pos, 90, 10, "Add Another Relationship", add_relationship_btn
			y_pos = y_pos + 15
			PushButton 55, y_pos, 125, 10, "Update Member Information", memb_info_change
			PushButton 410, y_pos, 60, 10, "NEXT MEMB", next_memb_btn

			Text 50, 370, 75, 10, "Principal Wage Earner: "
			DropListBox 130, 365, 125, 15, memb_droplist, the_pwe_for_this_case

			btn_pos = 70
			For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
				If i = memb_selected Then
					Text 9, btn_pos+1, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number
				Else
					PushButton 5, btn_pos, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number, HH_MEMB_ARRAY(i).button_one
				End If
				btn_pos = btn_pos + 10
			Next
			' PushButton 5, 70, 40, 10, "MEMB 01", memb_select_btn
			' PushButton 5, 80, 40, 10, "MEMB 01", Button9
			' PushButton 5, 90, 40, 10, "MEMB 01", Button10
			' PushButton 5, 100, 40, 10, "MEMB 01", Button11
			' PushButton 5, 110, 40, 10, "MEMB 01", Button12

			' PushButton 95, 10, 125, 15, "Review HH Member Information", HH_memb_detail_review
			Text 105, 13, 125, 15, "Review HH Member Information"
			PushButton 10, 10, 80, 15, "List of HH Members", hh_list_btn

		End If
		If page_display = show_q_1_2 Then
			Text 500, 42, 60, 13, "Q. 1 and 2"

			Text 15, 15, 450, 10, "Q. 1. Does everyone in your household buy, fix or eat food with you?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer

			Text 15, 50, 450, 10, "Q. 2. Is anyone who is in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			Text 135, 70, 65, 10, "Answer on the CAF"
			Text 265, 70, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 65, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 65, 125, 45, "", Combo2
		End If
		If page_display = show_q_3 Then
			Text 508, 57, 60, 13, "Q. 3"

			Text 15, 15, 450, 10, "Q. 3. Is anyone in your household attending school?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_4 Then
			Text 508, 72, 60, 13, "Q. 4"

			Text 15, 15, 450, 10, "Q. 4. Is anyone in your household temporarily not living in your home? (example: vacation, foster care, treatment, hospital job search)"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_5 Then
			Text 508, 87, 60, 13, "Q. 5"

			Text 15, 15, 450, 10, "Q. 5. Is anyone blind, or does anyone have a physical or mental health condition that limit the ability to work or perform daily activities?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_6 Then
			Text 508, 102, 60, 13, "Q. 6"

			Text 15, 15, 450, 10, "Q. 6. Is anyone unable to work for reasons other than illness or disability?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_7 Then
			Text 508, 117, 60, 13, "Q. 7"

			Text 15, 15, 450, 10, "Q. 7. In the last 60 days did anyone in the household: Stop working or quit? Refuse a job offer? Ask to work fewwer hours? Go on strike?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_8 Then
			Text 508, 132, 60, 13, "Q. 8"

			Text 15, 15, 450, 10, "Q. 8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer

			Text 15, 50, 450, 10, "Q. 8a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?"
			Text 135, 70, 65, 10, "Answer on the CAF"
			Text 265, 70, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 65, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 65, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_9 Then
			Text 508, 147, 60, 13, "Q. 9"

			Text 15, 15, 450, 20, "Q. 9. Does anyone in the household have a job or expect to get income from a job this month or next month? (Include income from Work Study and paid scholarships. Include free benefits or reduced expenses received for work (shelter, food, clothing, etc.)"
			Text 135, 40, 65, 10, "Answer on the CAF"
			Text 265, 40, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 35, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 35, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_10 Then
			Text 507, 162, 60, 13, "Q. 10"

			Text 15, 15, 450, 10, "Q. 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_11 Then
			Text 507, 177, 60, 13, "Q. 11"

			Text 15, 15, 450, 10, "Q. 11. Do you expect any changes in income, expenses or work hours?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_12 Then
			Text 507, 192, 60, 13, "Q. 12"

			Text 15, 15, 450, 10, "Q. 12. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_13 Then
			Text 507, 207, 60, 13, "Q. 13"

			Text 15, 15, 450, 10, "Q. 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_14_15 Then
			Text 495, 222, 60, 13, "Q. 14 and 15"

			Text 15, 15, 450, 10, "Q. 14. Does your household have the following housing expenses?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer

			Text 15, 50, 450, 10, "Q. 15. Does your household have the following utility expenses any time during the year?"
			Text 135, 70, 65, 10, "Answer on the CAF"
			Text 265, 70, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 65, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 65, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_16_18 Then
			Text 487, 237, 60, 13, "Q. 16, 17, and 18"

			Text 15, 15, 450, 20, "Q. 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs."
			Text 135, 40, 65, 10, "Answer on the CAF"
			Text 265, 40, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 35, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 35, 125, 45, "", confirm_caf_answer

			Text 15, 60, 450, 20, "Q. 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?"
			Text 135, 80, 65, 10, "Answer on the CAF"
			Text 265, 80, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 75, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 75, 125, 45, "", confirm_caf_answer

			Text 15, 100, 450, 20, "Q. 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?"
			Text 135, 120, 65, 10, "Answer on the CAF"
			Text 265, 120, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 115, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 115, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_19 Then
			Text 507, 252, 60, 13, "Q. 19"

			Text 15, 15, 450, 10, "Q. 19. For SNAP only: Does anyone in the household have medical expenses?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_20_21 Then
			Text 495, 267, 60, 13, "Q. 20 and 21"

			Text 15, 15, 450, 10, "Q. 20. Does anyone in the household own, or is anyone buying, any of the following?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer

			Text 15, 50, 450, 20, "Q. 21. FOR CASH PROGRAMS ONLY: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For Example: Cash, Bank Accounts, Stocks, Bonds, or Vehicles)?"
			Text 135, 75, 65, 10, "Answer on the CAF"
			Text 265, 75, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 70, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 70, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_22 Then
			Text 507, 282, 60, 13, "Q. 22"

			Text 15, 15, 450, 10, "Q. 22. FOR RECERTIFICATIONS ONLY: Did anyone move in or out of your home in the past 12 months?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_23 Then
			Text 507, 297, 60, 13, "Q. 23"

			Text 15, 15, 450, 10, "Q. 23. For children under the age of 19, are both parents living in the home?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_q_24 Then
			Text 507, 312, 60, 13, "Q. 24"

			Text 15, 15, 450, 10, "Q. 24. FOR MSA RECIPIENTS ONLY: Does anyone in the household have any of the following expenses?"
			Text 135, 35, 65, 10, "Answer on the CAF"
			Text 265, 35, 70, 10, "Confirm CAF Answer"
			DropListBox 200, 30, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", caf_answer
			ComboBox 340, 30, 125, 45, "", confirm_caf_answer
		End If
		If page_display = show_qual Then
			Text 492, 327, 60, 13, "CAF QUAL Q"
		End If
		If page_display = show_pg_last Then
			Text 490, 342, 60, 13, "CAF Last Page"
		End If

		If page_display <> show_pg_one Then PushButton 485, 10, 60, 13, "CAF Page 1", caf_page_one_btn
		If page_display <> show_pg_memb_list AND page_display <> show_pg_memb_info Then PushButton 485, 25, 60, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_q_1_2 Then PushButton 485, 40, 60, 13, "Q. 1 and 2", caf_q_1_2_btn
		If page_display <> show_q_3 Then PushButton 485, 55, 60, 13, "Q. 3", caf_q_3_btn
		If page_display <> show_q_4 Then PushButton 485, 70, 60, 13, "Q. 4", caf_q_4_btn
		If page_display <> show_q_5 Then PushButton 485, 85, 60, 13, "Q. 5", caf_q_5_btn
		If page_display <> show_q_6 Then PushButton 485, 100, 60, 13, "Q. 6", caf_q_6_btn
		If page_display <> show_q_7 Then PushButton 485, 115, 60, 13, "Q. 7", caf_q_7_btn
		If page_display <> show_q_8 Then PushButton 485, 130, 60, 13, "Q. 8", caf_q_8_btn
		If page_display <> show_q_9 Then PushButton 485, 145, 60, 13, "Q. 9", caf_q_9_btn
		If page_display <> show_q_10 Then PushButton 485, 160, 60, 13, "Q. 10", caf_q_10_btn
		If page_display <> show_q_11 Then PushButton 485, 175, 60, 13, "Q. 11", caf_q_11_btn
		If page_display <> show_q_12 Then PushButton 485, 190, 60, 13, "Q. 12", caf_q_12_btn
		If page_display <> show_q_13 Then PushButton 485, 205, 60, 13, "Q. 13", caf_q_13_btn
		If page_display <> show_q_14_15 Then PushButton 485, 220, 60, 13, "Q. 14 and 15", caf_q_14_15_btn
		If page_display <> show_q_16_18 Then PushButton 485, 235, 60, 13, "Q. 16, 17, and 18", caf_q_16_17_18_btn
		If page_display <> show_q_19 Then PushButton 485, 250, 60, 13, "Q. 19", caf_q_19_btn
		If page_display <> show_q_20_21 Then PushButton 485, 265, 60, 13, "Q. 20 and 21", caf_q_20_21_btn
		If page_display <> show_q_22 Then PushButton 485, 280, 60, 13, "Q. 22", caf_q_22_btn
		If page_display <> show_q_23 Then PushButton 485, 295, 60, 13, "Q. 23", caf_q_23_btn
		If page_display <> show_q_24 Then PushButton 485, 310, 60, 13, "Q. 24", caf_q_24_btn
		If page_display <> show_qual Then PushButton 485, 325, 60, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> show_pg_last Then PushButton 485, 340, 60, 13, "CAF Last Page", caf_last_page_btn
		PushButton 415, 365, 50, 15, "NEXT", next_btn
		PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn



	EndDialog
end function

'CAF PAGE 1 - HH Comp and Address
' FUNCTION - Read the current address on ADDR - resi and mail and phone numbers.
function read_ADDR_panel(addr_eff_date, line_one, line_two, city, state, zip, county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, phone_one, type_one, phone_two, type_two, phone_three, type_three, updated_date)
    Call navigate_to_MAXIS_screen("STAT", "ADDR")

    EMReadScreen line_one, 22, 6, 43
    EMReadScreen line_two, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 7, 9, 43
    EMReadScreen county_line, 2, 9, 66
    EMReadScreen verif_line, 2, 9, 74
    EMReadScreen homeless, 1, 10, 43
    EMReadScreen ind_reservation, 1, 10, 74
    EMReadScreen living_sit, 2, 11, 43
	EMReadScreen res_name, 2, 11, 74

    line_one = replace(line_one, "_", "")
    line_two = replace(line_two, "_", "")
    city = replace(city_line, "_", "")
    state = state_line
    zip = replace(zip_line, "_", "")

    If county_line = "01" Then county = "01 - Aitkin"
    If county_line = "02" Then county = "02 - Anoka"
    If county_line = "03" Then county = "03 - Becker"
    If county_line = "04" Then county = "04 - Beltrami"
    If county_line = "05" Then county = "05 - Benton"
    If county_line = "06" Then county = "06 - Big Stone"
    If county_line = "07" Then county = "07 - Blue Earth"
    If county_line = "08" Then county = "08 - Brown"
    If county_line = "09" Then county = "09 - Carlton"
    If county_line = "10" Then county = "10 - Carver"
    If county_line = "11" Then county = "11 - Cass"
    If county_line = "12" Then county = "12 - Chippewa"
    If county_line = "13" Then county = "13 - Chisago"
    If county_line = "14" Then county = "14 - Clay"
    If county_line = "15" Then county = "15 - Clearwater"
    If county_line = "16" Then county = "16 - Cook"
    If county_line = "17" Then county = "17 - Cottonwood"
    If county_line = "18" Then county = "18 - Crow Wing"
    If county_line = "19" Then county = "19 - Dakota"
    If county_line = "20" Then county = "20 - Dodge"
    If county_line = "21" Then county = "21 - Douglas"
    If county_line = "22" Then county = "22 - Faribault"
    If county_line = "23" Then county = "23 - Fillmore"
    If county_line = "24" Then county = "24 - Freeborn"
    If county_line = "25" Then county = "25 - Goodhue"
    If county_line = "26" Then county = "26 - Grant"
    If county_line = "27" Then county = "27 - Hennepin"
    If county_line = "28" Then county = "28 - Houston"
    If county_line = "29" Then county = "29 - Hubbard"
    If county_line = "30" Then county = "30 - Isanti"
    If county_line = "31" Then county = "31 - Itasca"
    If county_line = "32" Then county = "32 - Jackson"
    If county_line = "33" Then county = "33 - Kanabec"
    If county_line = "34" Then county = "34 - Kandiyohi"
    If county_line = "35" Then county = "35 - Kittson"
    If county_line = "36" Then county = "36 - Koochiching"
    If county_line = "37" Then county = "37 - Lac Qui Parle"
    If county_line = "38" Then county = "38 - Lake"
    If county_line = "39" Then county = "39 - Lake Of Woods"
    If county_line = "40" Then county = "40 - Le Sueur"
    If county_line = "41" Then county = "41 - Lincoln"
    If county_line = "42" Then county = "42 - Lyon"
    If county_line = "43" Then county = "43 - Mcleod"
    If county_line = "44" Then county = "44 - Mahnomen"
    If county_line = "45" Then county = "45 - Marshall"
    If county_line = "46" Then county = "46 - Martin"
    If county_line = "47" Then county = "47 - Meeker"
    If county_line = "48" Then county = "48 - Mille Lacs"
    If county_line = "49" Then county = "49 - Morrison"
    If county_line = "50" Then county = "50 - Mower"
    If county_line = "51" Then county = "51 - Murray"
    If county_line = "52" Then county = "52 - Nicollet"
    If county_line = "53" Then county = "53 - Nobles"
    If county_line = "54" Then county = "54 - Norman"
    If county_line = "55" Then county = "55 - Olmsted"
    If county_line = "56" Then county = "56 - Otter Tail"
    If county_line = "57" Then county = "57 - Pennington"
    If county_line = "58" Then county = "58 - Pine"
    If county_line = "59" Then county = "59 - Pipestone"
    If county_line = "60" Then county = "60 - Polk"
    If county_line = "61" Then county = "61 - Pope"
    If county_line = "62" Then county = "62 - Ramsey"
    If county_line = "63" Then county = "63 - Red Lake"
    If county_line = "64" Then county = "64 - Redwood"
    If county_line = "65" Then county = "65 - Renville"
    If county_line = "66" Then county = "66 - Rice"
    If county_line = "67" Then county = "67 - Rock"
    If county_line = "68" Then county = "68 - Roseau"
    If county_line = "69" Then county = "69 - St. Louis"
    If county_line = "70" Then county = "70 - Scott"
    If county_line = "71" Then county = "71 - Sherburne"
    If county_line = "72" Then county = "72 - Sibley"
    If county_line = "73" Then county = "73 - Stearns"
    If county_line = "74" Then county = "74 - Steele"
    If county_line = "75" Then county = "75 - Stevens"
    If county_line = "76" Then county = "76 - Swift"
    If county_line = "77" Then county = "77 - Todd"
    If county_line = "78" Then county = "78 - Traverse"
    If county_line = "79" Then county = "79 - Wabasha"
    If county_line = "80" Then county = "80 - Wadena"
    If county_line = "81" Then county = "81 - Waseca"
    If county_line = "82" Then county = "82 - Washington"
    If county_line = "83" Then county = "83 - Watonwan"
    If county_line = "84" Then county = "84 - Wilkin"
    If county_line = "85" Then county = "85 - Winona"
    If county_line = "86" Then county = "86 - Wright"
    If county_line = "87" Then county = "87 - Yellow Medicine"
    If county_line = "89" Then county = "89 - Out-of-State"

    If homeless = "Y" Then homeless = "Yes"
    If homeless = "N" Then homeless = "No"
    If ind_reservation = "Y" Then ind_reservation = "Yes"
    If ind_reservation = "N" Then ind_reservation = "No"

    If verif_line = "SF" Then verif = "SF - Shelter Form"
    If verif_line = "Co" Then verif = "CO - Coltrl Stmt"
    If verif_line = "MO" Then verif = "MO - Mortgage Papers"
    If verif_line = "TX" Then verif = "TX - Prop Tax Stmt"
    If verif_line = "CD" Then verif = "CD - Contrct for Deed"
    If verif_line = "UT" Then verif = "UT - Utility Stmt"
    If verif_line = "DL" Then verif = "DL - Driver Lic/State ID"
    If verif_line = "OT" Then verif = "OT - Other Document"
    If verif_line = "NO" Then verif = "NO - No Ver Prvd"
    If verif_line = "?_" Then verif = "? - Delayed"
    If verif_line = "__" Then verif = "Blank"


    If living_sit = "__" Then living_sit = "Blank"
    If living_sit = "01" Then living_sit = "01 - Own Housing (lease, mortgage, or roomate)"
    If living_sit = "02" Then living_sit = "02 - Family/Friends due to economic hardship"
    If living_sit = "03" Then living_sit = "03 - Servc prvdr- foster/group home"
    If living_sit = "04" Then living_sit = "04 - Hospital/Treatment/Detox/Nursing Home"
    If living_sit = "05" Then living_sit = "05 - Jail/Prison//Juvenile Det."
    If living_sit = "06" Then living_sit = "06 - Hotel/Motel"
    If living_sit = "07" Then living_sit = "07 - Emergency Shelter"
    If living_sit = "08" Then living_sit = "08 - Place not meant for Housing"
    If living_sit = "09" Then living_sit = "09 - Declined"
    If living_sit = "10" Then living_sit = "10 - Unknown"

	If res_name = "__" Then res_name = "Blank"
	If res_name = "BD" Then res_name = "Bois Forte - Deer Creek"
	If res_name = "BN" Then res_name = "Bois Forte - Nett Lake"
	If res_name = "BV" Then res_name = "Bois Forte - Vermillion Lk"
	If res_name = "FL" Then res_name = "Fond du Lac"
	If res_name = "GP" Then res_name = "Grand Portage"
	If res_name = "LL" Then res_name = "Leach Lake"
	If res_name = "LS" Then res_name = "Lower Sioux"
	If res_name = "ML" Then res_name = "Mille Lacs"
	If res_name = "PL" Then res_name = "Prairie Island Community"
	If res_name = "RL" Then res_name = "Red Lake"
	If res_name = "SM" Then res_name = "Shakopee Mdewakanton"
	If res_name = "US" Then res_name = "Upper Sioux"
	If res_name = "WE" Then res_name = "White Earth"

    EMReadScreen addr_eff_date, 8, 4, 43
    EMReadScreen addr_future_date, 8, 4, 66
    EMReadScreen mail_line_one, 22, 13, 43
    EMReadScreen mail_line_two, 22, 14, 43
    EMReadScreen mail_city, 15, 15, 43
    EMReadScreen mail_state, 2, 16, 43
    EMReadScreen mail_zip, 7, 16, 52

    addr_eff_date = replace(addr_eff_date, " ", "/")
    addr_future_date = trim(addr_future_date)
    addr_future_date = replace(addr_future_date, " ", "/")
    mail_line_one = replace(mail_line_one, "_", "")
    mail_line_two = replace(mail_line_two, "_", "")
    mail_city = replace(mail_city, "_", "")
    mail_state = replace(mail_state, "_", "")
    mail_zip = replace(mail_zip, "_", "")

	EMReadScreen phone_one, 14, 17, 45
	EMReadScreen phone_two, 14, 18, 45
	EMReadScreen phone_three, 14, 19, 45
	EMReadScreen type_one, 1, 17, 67
	EMReadScreen type_two, 1, 18, 67
	EMReadScreen type_three, 1, 19, 67

	phone_one = "(" & replace(replace(replace(phone_one, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_one = "(___) ___ - ____" Then phone_one = ""
	If type_one = "_" Then type_one = "Unknown"
	If type_one = "H" Then type_one = "Home"
	If type_one = "W" Then type_one = "Work"
	If type_one = "C" Then type_one = "Cell"
	If type_one = "M" Then type_one = "Message"
	If type_one = "T" Then type_one = "TTY/TDD"

	phone_two = "(" & replace(replace(replace(phone_two, " ) ", ") "), " ", " - "), ")", ") ")
	If phone_two = "(___) ___ - ____" Then phone_two = ""
	If type_two = "_" Then type_two = "Unknown"
	If type_two = "H" Then type_two = "Home"
	If type_two = "W" Then type_two = "Work"
	If type_two = "C" Then type_two = "Cell"
	If type_two = "M" Then type_two = "Message"
	If type_two = "T" Then type_two = "TTY/TDD"

	phone_three = "(" & replace(replace(replace(phone_three, " ) ", ") "), " ", " - "), ")", ") ")
	If phone_three = "(___) ___ - ____" Then phone_three = ""
	If type_three = "_" Then type_three = "Unknown"
	If type_three = "H" Then type_three = "Home"
	If type_three = "W" Then type_three = "Work"
	If type_three = "C" Then type_three = "Cell"
	If type_three = "M" Then type_three = "Message"
	If type_three = "T" Then type_three = "TTY/TDD"

	EMReadScreen updated_date, 8, 21, 55
	updated_date = replace(updated_date, " ", "/")
end function

' FUNCTION - Read all the HH Members from case, including how they are related to M01 and SIBL and PARE - get age
function read_all_the_MEMBs()


	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 2, 4, 33
		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
		End If
		If client_array <> "" Then client_array = client_array & "|" & ref_nbr
		If client_array = "" Then client_array = client_array & ref_nbr
		transmit      'Going to the next MEMB panel
		Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
		member_count = member_count + 1
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
	' MsgBox client_array
	client_array = split(client_array, "|")

	clt_count = 0

	For each hh_clt in client_array

		ReDim Preserve HH_MEMB_ARRAY(clt_count)
		Set HH_MEMB_ARRAY(clt_count) = new mx_hh_member
		HH_MEMB_ARRAY(clt_count).ref_number = hh_clt
		HH_MEMB_ARRAY(clt_count).define_the_member
		HH_MEMB_ARRAY(clt_count).button_one = 500 + clt_count
		memb_droplist = memb_droplist+chr(9)+HH_MEMB_ARRAY(clt_count).ref_number & " - " & HH_MEMB_ARRAY(clt_count).full_name
		If HH_MEMB_ARRAY(clt_count).fs_pwe = "Yes" Then the_pwe_for_this_case = HH_MEMB_ARRAY(clt_count).ref_number & " - " & HH_MEMB_ARRAY(clt_count).full_name

		clt_count = clt_count + 1
	Next

	rela_counter = 0
	For i = 0 to UBOUND(HH_MEMB_ARRAY, 1)
		If HH_MEMB_ARRAY(i).rel_to_applcnt <> "Self" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Not Related" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Live-in Attendant" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Unknown" Then
			ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = "01"
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(0).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = HH_MEMB_ARRAY(i).rel_to_applcnt

			rela_counter = rela_counter + 1

			ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

		 	' MsgBox "Member Count - " & i & vbNewLine & "Relationship - " & HH_MEMB_ARRAY(i).rel_to_applcnt
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = "01"
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(0).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Spouse" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Spouse"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Child" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Parent"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Parent" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Child"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Sibling" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Sibling"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Step Sibling" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Sibling"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Step Child" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Parent"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Step Parent" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Child"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Aunt" OR HH_MEMB_ARRAY(i).rel_to_applcnt = "Uncle" Then
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Neice"
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Nephew"
			End If
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Nephew" OR HH_MEMB_ARRAY(i).rel_to_applcnt = "Neice" Then
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Aunt"
				If HH_MEMB_ARRAY(0).gender = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Uncle"
			End If
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Cousin" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Cousin"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Grandparent" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandchild"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Grandchild" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandparent"
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Other Relative" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Other Relative"

			rela_counter = rela_counter + 1
		End If
	Next

	Call navigate_to_MAXIS_screen("STAT", "SIBL")

	sibl_row = 7
	Do
		EMReadScreen sibl_group_nbr, 2, sibl_row, 28
		If sibl_group_nbr <> "__" Then
			sibl_col = 39
			Do
				EMReadScreen sibl_ref_nrb, 2, sibl_row, sibl_col
				If sibl_ref_nrb <> "__" Then
					list_of_siblings = list_of_siblings & "~" & sibl_ref_nrb
					sibl_col = sibl_col + 4
				End If
			Loop until sibl_ref_nrb = "__"
			list_of_siblings = right(list_of_siblings, len(list_of_siblings) - 1)
			sibl_array = split(list_of_siblings, "~")
			For each memb_sibling in sibl_array
				For each other_sibling in sibl_array
					If memb_sibling <> other_sibling Then
						ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

						ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = memb_sibling
						ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = other_sibling
						ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Sibling"

						For i = 0 to UBOUND(HH_MEMB_ARRAY, 1)
							If HH_MEMB_ARRAY(i).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
							If HH_MEMB_ARRAY(i).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
						Next

						rela_counter = rela_counter + 1
					End If
				Next
			Next
		End If
		sibl_row = sibl_row + 1
	Loop until sibl_group_nbr = "__"

	Call navigate_to_MAXIS_screen("STAT", "PARE")

	'Need to add a way to find relationship verifications for MEMB 01
	EMWriteScreen HH_MEMB_ARRAY(0).ref_number, 20, 76
	transmit



	For i = 0 to UBound(HH_MEMB_ARRAY, 1)						'we start with 1 because 0 is MEMB 01 and that parental relationshipare all known because of MEMB
		EMWriteScreen HH_MEMB_ARRAY(i).ref_number, 20, 76
		transmit

		pare_row = 8
		Do
			EMReadScreen child_ref_nbr, 2, pare_row, 24
			If child_ref_nbr <> "__" Then
				If i = 0 Then
					For known_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
						If ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, known_rela) = "01" AND child_ref_nbr = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, known_rela) THen
							EMReadScreen pare_verif, 2, pare_row, 71
							If pare_verif = "BC" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "BC - Birth Certificate"
							If pare_verif = "AR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "AR - Adoption Records"
							If pare_verif = "LG" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "LG = Legal Guardian"
							If pare_verif = "RE" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "RE - Religious Records"
							If pare_verif = "HR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "HR - Hospital Records"
							If pare_verif = "RP" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "RP - Recognition of Parentage"
							If pare_verif = "OT" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "OT - Other Verifciation"
							If pare_verif = "NO" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, known_rela) = "NO - No Verif Provided"
						End If
					Next
				Else
					EMReadScreen pare_type, 1, pare_row, 53
					EMReadScreen pare_verif, 2, pare_row, 71

					ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = child_ref_nbr

					If pare_type = "1" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Parent"
					If pare_type = "2" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Parent"
					If pare_type = "3" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandparent"
					If pare_type = "4" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Relative Caregiver"
					If pare_type = "5" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Foster Parent"
					If pare_type = "6" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Not Related"
					If pare_type = "7" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Legal Guardian"
					If pare_type = "8" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Other Relative"

					If pare_verif = "BC" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "BC - Birth Certificate"
					If pare_verif = "AR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "AR - Adoption Records"
					If pare_verif = "LG" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "LG = Legal Guardian"
					If pare_verif = "RE" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RE - Religious Records"
					If pare_verif = "HR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "HR - Hospital Records"
					If pare_verif = "RP" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RP - Recognition of Parentage"
					If pare_verif = "OT" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "OT - Other Verifciation"
					If pare_verif = "NO" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "NO - No Verif Provided"

					for x = 0 to UBound(HH_MEMB_ARRAY, 1)
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
					Next

					rela_counter = rela_counter + 1

					ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = child_ref_nbr
					ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number

					If pare_type = "1" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Child"
					If pare_type = "2" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Step Child"
					If pare_type = "3" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Grandchild"
					If pare_type = "4" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Relative Caregiver"
					If pare_type = "5" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Foster Child"
					If pare_type = "6" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Not Related"
					If pare_type = "7" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Legal Guardian"
					If pare_type = "8" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Other Relative"

					If pare_verif = "BC" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "BC - Birth Certificate"
					If pare_verif = "AR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "AR - Adoption Records"
					If pare_verif = "LG" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "LG = Legal Guardian"
					If pare_verif = "RE" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RE - Religious Records"
					If pare_verif = "HR" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "HR - Hospital Records"
					If pare_verif = "RP" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "RP - Recognition of Parentage"
					If pare_verif = "OT" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "OT - Other Verifciation"
					If pare_verif = "NO" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, rela_counter) = "NO - No Verif Provided"

					for x = 0 to UBound(HH_MEMB_ARRAY, 1)
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
						If HH_MEMB_ARRAY(x).ref_number = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) Then ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(x).full_name
					Next

					rela_counter = rela_counter + 1

				End If
			End If
			pare_row = pare_row + 1

		Loop until child_ref_nbr = "__"
	Next

	For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
		ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_one, the_rela) = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) & " - " & ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, the_rela)
		ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela) = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, the_rela) & " - " & ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, the_rela)

		' MsgBox "Relationship detail:" & vbNewLine & ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_one, the_rela) & " is the " & ALL_HH_RELATIONSHIPS_ARRAY(rela_type, the_rela) & " of " & ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela)
	Next
end function

' FUNCTION - have worker enter the ADDR that is on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmed this address.
' FUNCTION - have worker enter the mailing address on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmaed this address. Add functionality for General Deliver and explaining the requirements
' FUNCTION - list all of the phone numbers from ADDR and have the worker indicate in droplist if they are correct and the type of phone number that it is. If none - then have the worker clarify if there is another phone number or not
function review_ADDR_information()
end function

' FUNCTION - Ask about living situation - worker to select from CAF and verbally confirm. A;sp ask if HOMELESS - if yes then - want to speak to shelter team - do you lack access to work related necessities
function review_living_situation()
end function

' FUNCTION - List all of the HH Members and their relationships. Ask if all of these people still live in the house - Ask if there is any other person living at this address - Confirm the relationships - add new people if needed
function review_the_MEMBs()
end function

' FUNCTION - Lopp through all of the HH Members and confirm:
	' Confirm personal information (name, DOB, marital status, SSN - if not validated)
	' Detail if ID is needed and if we have it
	' Confirm citizenship/immigration
	' Determine language
	' Determine if recently arrived in MN
	' Determine relationship to other HH Members - determine proof if needed
	' Determine who Principal Wage Earner is
function review_MEMB_detail()
end function

' FUNCTION - Get AREP information or confirm no AREP
' FUNCTION - Get SWKR information
function review_associeated_people()
end function


'CAF QUESIION

' 1. Does everyone in your household buy, fix, or eat food with you?
' 2. Is anyone who is in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?

'FUNCTION - CAF Q1 & Q2 Enter what CAF has listed and then the verbal confirmation
function reveiw_caf_q_1_2()
end function

' 3. Is anyone in your household attending school?

'FUNCTION - CAF Q3 Enter the CAF answer and the verbal response - if NO move on. If YES then list all members and indicate what kind of school they are attending.
function review_caf_q_3()
end function

' 4. Is anyone in your household temporarily not living in your home? (example: vacation, foster care, treatment, hospital job search)

'FUNCTION - CAF Q4 enter the CAF answer and the verbal response. If No - move on - if yes the who and where they are - explain temp absense
function review_caf_q_4()
end function

' 5.	Is anyone blind, or does anyone have a physical or mental health condition that limit the ability to work or perform daily activities?

' 'FUNCTION - read MAXIS for DISA panels
' function
' end function

'FUNCTION - CAF Q5 enter the CAF answer and the verbal answer. If no - next question, if yes - then WHO and what verifs or program details
function review_caf_q_5()
end function

' 'FUNCTIOn - review the panels and update/delete as needed.
' function
' end function

'6.	Is anyone unable to work for reasons other than illness or disability?

'FUNCTION - CAF Q6 enter CAF answer and verbal response. If no - keep going. If yes - list HH members and the reason for inability to work - this might be important for ABAWD or MFIP extension - EXPAND ON POSSIBLE ABAWD EXEMPTIONS '
function review_caf_q_6()
end function

'7.	In the last 60 days did anyone in the household: Stop working or quit? Refuse a job offer? Ask to work fewwer hours? Go on strike?

'FUNCTION - CAF Q7 enter CAF answer and verbal response. If no - move on. If yes - who, and then details based on program.
function review_caf_q_7()
end function

'8.	Has anyone in the household had a job or been self-employed in the past 12 months?

'FUNCTION - CAF Q8 - enter CAF response and verbal response. If yes gather the details. If SNAP ask about 36 months - CAF response and verbal response - if yes - gather detail'
function review_caf_q_8()
end function

'9.	Does anyone in the household have a job or expect to get income from a job this month or next month? (Include income from Work Study and paid scholarships. Include free benefits or reduced expenses received for work (shelter, food, clothing, etc.)

'FUNCTION - CAF Q9 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each job - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.
function review_caf_q_9()
end function

'10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?

'FUNCTION - CAF Q10 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each self employment - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.
function review_caf_q_10()
end function

'11. Do you expect any changes in income, expenses or work hours?

'FUNCTION - CAF Q11 - enter the CAF response and verbal response - If yes then list all the income and have them enter which is changing, when and how.
function review_caf_q_11()
end function

'12. Has anyone in the household applied for or does anyone get any of the following types of income each month?

'FUNCTION - CAF Q12 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each income source or applicaiton - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.
function review_caf_q_12()
end function

'13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?

'FUNCTION - CAF Q13 - enter the CAF response and the verbal response. If yes then list the persons and use the school information gathered before to indicate who might have this and what student income there is
function review_caf_q_13()
end function

'14. Does your household have the following housing expenses?
'15. Does your household have the following utility expenses any time during the year?

'FUNCTION - CAF Q14 and Q15 - list all of the expenses with the CAF response and the verbal response then the detail.
function review_caf_q_14_15()
end function

'16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs.
'17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?
'18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?

'FUNCTION - CAF Q16 & Q17 & Q18 -  enter the CAF answers and verbal responses. If yes to either - gather details about who, what, how much, and etc
function review_caf_q_16_17_18()
end function

'19. For SNAP only: Does anyone in the household have medical expenses?

'FUNCTION - CAF Q19 - SNAP only - CAF response and verbal response gather detail if Yes
function review_caf_q_19()
end function

'20. Does anyone in the household own, or is anyone buying, any of the following?
'21. FOR CASH PROGRAMS ONLY: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For Example: Cash, Bank Accounts, Stocks, Bonds, or Vehicles)?

'FUNCTION - CAF Q20 & Q21 - All the asset information - for cash programs we need some verif - for SNAP if at application - enter balances - no verif - EXP
function review_caf_q_20_21()
end function

'22. FOR RECERTIFICATIONS ONLY: Did anyone move in or out of your home in the past 12 months?

'FUNCTION - CAF Q22 - Enter CAF response and verbal response
function review_caf_q_22()
end function

'23. For children under the age of 19, are both parents living in the home?

'FUNCTION - CAF Q23 - Enter CAF response and verbal response
function review_caf_q_23()
end function

'24. FOR MSA RECIPIENTS ONLY: Does anyone in the household have any of the following expenses?

'FUNCTION - CAF Q24 - Enter CAF response and verbal response
function review_caf_q_24()
end function

'25. Penalty warning and qualification questions

'FUNCTION - CAF Q25 - Enter CAF response and verbal response
function review_caf_q_25()
end function

'26. Did client sign the last page of the CAF?

'FUNCTION - CAF Q26 - Enter CAF response and verbal response
function review_caf_q_26()
end function



'THE SCRIPT ===========================================================================================================
EMConnect ""
Call check_for_MAXIS(TRUE)

Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
interview_date = date & ""

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
End If
'
'Start of Interview
'Dialog to gather interview set up information/detail.
'Date of interview, type of interview, case number.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 105, "INTERVIEW Case number dialog"
  EditBox 65, 10, 60, 15, MAXIS_case_number
  EditBox 210, 10, 15, 15, MAXIS_footer_month
  EditBox 230, 10, 15, 15, MAXIS_footer_year
  CheckBox 10, 45, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 50, 45, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 90, 45, 35, 10, "EMER", EMER_on_CAF_checkbox
  DropListBox 135, 45, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"ApplyMN"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  EditBox 60, 65, 45, 15, interview_date
  ComboBox 165, 65, 110, 45, "Type or Select"+chr(9)+"phone"+chr(9)+"office", interview_type
  ButtonGroup ButtonPressed
    PushButton 35, 85, 15, 15, "!", tips_and_tricks_button
    OkButton 170, 85, 50, 15
    CancelButton 225, 85, 50, 15
  GroupBox 5, 30, 125, 30, "Programs marked on CAF"
  Text 135, 35, 65, 10, "Actual CAF Form:"
  Text 55, 90, 105, 10, "Look for me for Tips and Tricks!"
  Text 10, 70, 50, 10, "Interview Date:"
  Text 115, 70, 50, 10, "Interview Type:"
  Text 140, 15, 65, 10, "Footer month/year: "
  Text 10, 15, 50, 10, "Case number:"
EndDialog

Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_without_confirmation

		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call read_ADDR_panel(addr_eff_date, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, phone_numb_one, phone_type_one, phone_numb_two, phone_type_two, phone_numb_three, phone_type_three, addr_updated_date)
updated_addr_eff_date 		= addr_eff_date
updated_resi_line_one 		= resi_line_one
updated_resi_line_two 		= resi_line_two
updated_resi_city 			= resi_city
updated_resi_state 			= resi_state
updated_resi_zip 			= resi_zip
updated_resi_county 		= resi_county
updated_verif 				= verif
updated_homeless 			= homeless
updated_ind_reservation 	= ind_reservation
updated_living_sit 			= living_sit
updated_res_name 			= res_name
updated_mail_line_one 		= mail_line_one
updated_mail_line_two 		= mail_line_two
updated_mail_city 			= mail_city
updated_mail_state 			= mail_state
updated_mail_zip 			= mail_zip
updated_phone_numb_one 		= phone_numb_one
updated_phone_type_one 		= phone_type_one
updated_phone_numb_two 		= phone_numb_two
updated_phone_type_two 		= phone_type_two
updated_phone_numb_three 	= phone_numb_three
updated_phone_type_three 	= phone_type_three

Call read_all_the_MEMBs

For hh_memb = 0 to UBound(HH_MEMB_ARRAY, 1)
	' MsgBox HH_MEMB_ARRAY(hh_memb).ref_number & vbNewLine & HH_MEMB_ARRAY(hh_memb).full_name
Next
'Gather all of the household information

'Go to see if recert or application and CAF date
'Gather all of the HH Member information
'Gather the DISA information

'Dialog to get the programs, person the interview is completed with, and some basic CAF information
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 326, 45, "Interview Details"
  EditBox 50, 5, 50, 15, caf_date
  ComboBox 175, 5, 145, 45, "Type or Select"+chr(9)+"List of People", interview_with
  ButtonGroup ButtonPressed
    OkButton 220, 25, 50, 15
    CancelButton 270, 25, 50, 15
  Text 10, 10, 35, 10, "CAF Date:"
  Text 120, 10, 50, 10, "Interview With:"
EndDialog
Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_without_confirmation

		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Button Definitions
caf_page_one_btn	= 1000
caf_membs_btn		= 1001
caf_q_1_2_btn		= 1002
caf_q_3_btn			= 1003
af_q_4_btn 			= 1004
caf_q_5_btn			= 1005
caf_q_6_btn			= 1006
caf_q_7_btn			= 1007
caf_q_8_btn			= 1008
caf_q_9_btn			= 1009
caf_q_10_btn		= 1010
caf_q_11_btn		= 1011
caf_q_12_btn		= 1012
caf_q_13_btn		= 1013
caf_q_14_15_btn		= 1014
caf_q_16_17_18_btn	= 1015
caf_q_19_btn		= 1016
caf_q_20_21_btn		= 1017
caf_q_22_btn		= 1018
caf_q_23_btn		= 1019
caf_q_24_btn		= 1020
caf_qual_q_btn		= 1021
caf_last_page_btn	= 1022
next_btn			= 1023
finish_interview_btn= 1024

client_verbal_changes_resi_address_btn	= 2000
client_verbal_changes_mail_address_btn	= 2001
caf_info_different_btn					= 2002
client_verbal_changes_phone_numbers_btn	= 2003
add_memb_to_list_btn					= 2004
HH_memb_detail_review					= 2005
add_relationship_btn					= 2006
memb_info_change						= 2007
next_memb_btn							= 2008
hh_list_btn								= 2009

done_pg_one 	= FALSE
done_pg_memb 	= FALSE
done_q_1_2		= FALSE
done_q_3		= FALSE
done_q_4		= FALSE
done_q_5		= FALSE
done_q_6		= FALSE
done_q_7		= FALSE
done_q_8		= FALSE
done_q_9		= FALSE
done_q_10		= FALSE
done_q_11		= FALSE
done_q_12		= FALSE
done_q_13		= FALSE
done_q_14_15	= FALSE
done_q_16_18	= FALSE
done_q_19		= FALSE
done_q_20_21 	= FALSE
done_q_22		= FALSE
done_q_23		= FALSE
done_q_24		= FALSE
done_qual		= FALSE
done_pg_last	= FALSE


page_display = 1
show_pg_one 		= 1
show_pg_memb_list 	= 2
show_pg_memb_info 	= 3
show_q_1_2			= 4
show_q_3			= 5
show_q_4			= 6
show_q_5			= 7
show_q_6			= 8
show_q_7			= 9
show_q_8			= 10
show_q_9			= 11
show_q_10			= 12
show_q_11			= 13
show_q_12			= 14
show_q_13			= 15
show_q_14_15		= 16
show_q_16_18		= 17
show_q_19			= 18
show_q_20_21 		= 19
show_q_22			= 20
show_q_23			= 21
show_q_24			= 22
show_qual			= 23
show_pg_last		= 24

ButtonPressed = caf_page_one_btn
leave_loop = FALSE
Do
	Do
		If memb_selected = "" Then memb_selected = 0
		' MsgBox page_display

		call define_main_dialog

		err_msg = ""

		prev_page = page_display


		dialog Dialog1
		cancel_confirmation

		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

		If page_display <> prev_page Then
			'ADD FUNCTIONS HERE TO EVALUATE THE COMPLETION OF EACH PAGE
		End If

		' MsgBox "ButtonPressed - " & ButtonPressed
		For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
			' MsgBox HH_MEMB_ARRAY(i).button_one
			If ButtonPressed = HH_MEMB_ARRAY(i).button_one Then
				' MsgBox "selected"
				memb_selected = i
			End If
		Next
		If page_display = show_pg_memb_info AND ButtonPressed = -1 Then ButtonPressed = next_memb_btn
		If ButtonPressed = next_memb_btn Then
			memb_selected = memb_selected + 1
			If memb_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
		End If
		If ButtonPressed = -1 Then ButtonPressed = next_btn
		If ButtonPressed = next_btn Then
			If page_display = show_pg_one Then ButtonPressed = caf_membs_btn
			If page_display = show_pg_memb_list Then ButtonPressed = HH_memb_detail_review
			If page_display = show_pg_memb_info Then ButtonPressed = caf_q_1_2_btn
			If page_display = show_q_1_2 Then ButtonPressed = caf_q_3_btn
			If page_display = show_q_3 Then ButtonPressed = caf_q_4_btn
			If page_display = show_q_4 Then ButtonPressed = caf_q_5_btn
			If page_display = show_q_5 Then ButtonPressed = caf_q_6_btn
			If page_display = show_q_6 Then ButtonPressed = caf_q_7_btn
			If page_display = show_q_7 Then ButtonPressed = caf_q_8_btn
			If page_display = show_q_8 Then ButtonPressed = caf_q_9_btn
			If page_display = show_q_9 Then ButtonPressed = caf_q_10_btn
			If page_display = show_q_10 Then ButtonPressed = caf_q_11_btn
			If page_display = show_q_11 Then ButtonPressed = caf_q_12_btn
			If page_display = show_q_12 Then ButtonPressed = caf_q_13_btn
			If page_display = show_q_13 Then ButtonPressed = caf_q_14_15_btn
			If page_display = show_q_14_15 Then ButtonPressed = caf_q_16_17_18_btn
			If page_display = show_q_16_18 Then ButtonPressed = caf_q_19_btn
			If page_display = show_q_19 Then ButtonPressed = caf_q_20_21_btn
			If page_display = show_q_20_21 Then ButtonPressed = caf_q_22_btn
			If page_display = show_q_22 Then ButtonPressed = caf_q_23_btn
			If page_display = show_q_23 Then ButtonPressed = caf_q_24_btn
			If page_display = show_q_24 Then ButtonPressed = caf_qual_q_btn
			If page_display = show_qual Then ButtonPressed = caf_last_page_btn
			' If page_display = show_pg_last Then ButtonPressed =

		End If

		If ButtonPressed = caf_page_one_btn Then
			page_display = show_pg_one
		End If
		If ButtonPressed = caf_membs_btn Then
			page_display = show_pg_memb_list
		End If
		If ButtonPressed = hh_list_btn Then
			page_display = show_pg_memb_list
		End If
		If ButtonPressed = HH_memb_detail_review Then
			page_display = show_pg_memb_info
		End If
		If ButtonPressed = caf_q_1_2_btn Then
			page_display = show_q_1_2
		End If
		If ButtonPressed = caf_q_3_btn Then
			page_display = show_q_3
		End If
		If ButtonPressed = caf_q_4_btn Then
			page_display = show_q_4
		End If
		If ButtonPressed = caf_q_5_btn Then
			page_display = show_q_5
		End If
		If ButtonPressed = caf_q_6_btn Then
			page_display = show_q_6
		End If
		If ButtonPressed = caf_q_7_btn Then
			page_display = show_q_7
		End If
		If ButtonPressed = caf_q_8_btn Then
			page_display = show_q_8
		End If
		If ButtonPressed = caf_q_9_btn Then
			page_display = show_q_9
		End If
		If ButtonPressed = caf_q_10_btn Then
			page_display = show_q_10
		End If
		If ButtonPressed = caf_q_11_btn Then
			page_display = show_q_11
		End If
		If ButtonPressed = caf_q_12_btn Then
			page_display = show_q_12
		End If
		If ButtonPressed = caf_q_13_btn Then
			page_display = show_q_13
		End If
		If ButtonPressed = caf_q_14_15_btn Then
			page_display = show_q_14_15
		End If
		If ButtonPressed = caf_q_16_17_18_btn Then
			page_display = show_q_16_18
		End If
		If ButtonPressed = caf_q_19_btn Then
			page_display = show_q_19
		End If
		If ButtonPressed = caf_q_20_21_btn Then
			page_display = show_q_20_21
		End If
		If ButtonPressed = caf_q_22_btn Then
			page_display = show_q_22
		End If
		If ButtonPressed = caf_q_23_btn Then
			page_display = show_q_23
		End If
		If ButtonPressed = caf_q_24_btn Then
			page_display = show_q_24
		End If
		If ButtonPressed = caf_qual_q_btn Then
			page_display = show_qual
		End If
		If ButtonPressed = caf_last_page_btn Then
			page_display = show_pg_last
		End If

	Loop until leave_loop = TRUE
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


'CAF PAGE 1 - HH Comp and Address
' FUNCTION - Read the current address on ADDR - resi and mail and phone numbers.
' FUNCTION - have worker enter the ADDR that is on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmed this address.
' FUNCTION - have worker enter the mailing address on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmaed this address. Add functionality for General Deliver and explaining the requirements
' FUNCTION - list all of the phone numbers from ADDR and have the worker indicate in droplist if they are correct and the type of phone number that it is. If none - then have the worker clarify if there is another phone number or not

' FUNCTION - Ask about living situation - worker to select from CAF and verbally confirm. A;sp ask if HOMELESS - if yes then - want to speak to shelter team - do you lack access to work related necessities
' FUNCTION - Read all the HH Members from case, including how they are related to M01 and SIBL and PARE - get age
' FUNCTION - List all of the HH Members and their relationships. Ask if all of these people still live in the house - Ask if there is any other person living at this address - Confirm the relationships
' FUNCTION - Lopp through all of the HH Members and confirm:
	' Confirm personal information (name, DOB, marital status, SSN - if not validated)
	' Detail if ID is needed and if we have it
	' Confirm citizenship/immigration
	' Determine language
	' Determine if recently arrived in MN
	' Determine relationship to other HH Members - determine proof if needed
	' Determine who Principal Wage Earner is
' FUNCTION - Get AREP information or confirm no AREP
' FUNCTION - Get SWKR information


'CAF QUESIION

' 1. Does everyone in your household buy, fix, or eat food with you?
' 2. Is anyone who is in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?

'FUNCTION - CAF Q1 & Q2 Enter what CAF has listed and then the verbal confirmation

' 3. Is anyone in your household attending school?

'FUNCTION - CAF Q3 Enter the CAF answer and the verbal response - if NO move on. If YES then list all members and indicate what kind of school they are attending.

' 4. Is anyone in your household temporarily not living in your home? (example: vacation, foster care, treatment, hospital job search)

'FUNCTION - CAF Q4 enter the CAF answer and the verbal response. If No - move on - if yes the who and where they are - explain temp absense

' 5.	Is anyone blind, or does anyone have a physical or mental health condition that limit the ability to work or perform daily activities?

'FUNCTION - read MAXIS for DISA panels
'FUNCTION - CAF Q5 enter the CAF answer and the verbal answer. If no - next question, if yes - then WHO and what verifs or program details
'FUNCTIOn - review the panels and update/delete as needed.

'6.	Is anyone unable to work for reasons other than illness or disability?

'FUNCTION - CAF Q6 enter CAF answer and verbal response. If no - keep going. If yes - list HH members and the reason for inability to work - this might be important for ABAWD or MFIP extension - EXPAND ON POSSIBLE ABAWD EXEMPTIONS '

'7.	In the last 60 days did anyone in the household: Stop working or quit? Refuse a job offer? Ask to work fewwer hours? Go on strike?

'FUNCTION - CAF Q7 enter CAF answer and verbal response. If no - move on. If yes - who, and then details based on program.

'8.	Has anyone in the household had a job or been self-employed in the past 12 months?

'FUNCTION - CAF Q8 - enter CAF response and verbal response. If yes gather the details. If SNAP ask about 36 months - CAF response and verbal response - if yes - gather detail'

'9.	Does anyone in the household have a job or expect to get income from a job this month or next month? (Include income from Work Study and paid scholarships. Include free benefits or reduced expenses received for work (shelter, food, clothing, etc.)

'FUNCTION - CAF Q9 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each job - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.

'10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?

'FUNCTION - CAF Q10 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each self employment - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.

'11. Do you expect any changes in income, expenses or work hours?

'FUNCTION - CAF Q11 - enter the CAF response and verbal response - If yes then list all the income and have them enter which is changing, when and how.

'12. Has anyone in the household applied for or does anyone get any of the following types of income each month?

'FUNCTION - CAF Q12 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each income source or applicaiton - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.

'13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?

'FUNCTION - CAF Q13 - enter the CAF response and the verbal response. If yes then list the persons and use the school information gathered before to indicate who might have this and what student income there is

'14. Does your household have the following housing expenses?
'15. Does your household have the following utility expenses any time during the year?

'FUNCTION - CAF Q14 and Q15 - list all of the expenses with the CAF response and the verbal response then the detail.

'16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs.
'17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?
'18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?

'FUNCTION - CAF Q16 & Q17 & Q18 -  enter the CAF answers and verbal responses. If yes to either - gather details about who, what, how much, and etc

'19. For SNAP only: Does anyone in the household have medical expenses?

'FUNCTION - CAF Q19 - SNAP only - CAF response and verbal response gather detail if Yes

'20. Does anyone in the household own, or is anyone buying, any of the following?
'21. FOR CASH PROGRAMS ONLY: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For Example: Cash, Bank Accounts, Stocks, Bonds, or Vehicles)?

'FUNCTION - CAF Q20 & Q21 - All the asset information - for cash programs we need some verif - for SNAP if at application - enter balances - no verif - EXP

'22. FOR RECERTIFICATIONS ONLY: Did anyone move in or out of your home in the past 12 months?

'FUNCTION - CAF Q22 - Enter CAF response and verbal response

'23. For children under the age of 19, are both parents living in the home?

'FUNCTION - CAF Q23 - Enter CAF response and verbal response

'24. FOR MSA RECIPIENTS ONLY: Does anyone in the household have any of the following expenses?

'FUNCTION - CAF Q24 - Enter CAF response and verbal response

'25. Penalty warning and qualification questions

'FUNCTION - CAF Q25 - Enter CAF response and verbal response

'26. Did client sign the last page of the CAF?

'FUNCTION - CAF Q26 - Enter CAF response and verbal response
